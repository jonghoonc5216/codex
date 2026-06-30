from __future__ import annotations

import argparse
import html
import io
import json
import math
import mimetypes
import re
import subprocess
import sys
import time
import urllib.request
import zipfile
from dataclasses import dataclass
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

from dxf_parser import (
    KOREA_CRS,
    build_project,
    features_to_kml,
    lonlat_to_projected,
    projected_to_lonlat,
    resolve_image_path,
)


BUNDLE_ROOT = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
ROOT = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
STATIC_ROOT = BUNDLE_ROOT / "static"
DEFAULT_DXF = BUNDLE_ROOT / "data" / "sample_yulha.dxf"
UPLOAD_ROOT = ROOT / "data" / "uploads"
TILE_CACHE_ROOT = ROOT / "cache" / "vworld-satellite"
EXPORT_ROOT = ROOT / "exports"
DEFAULT_SOURCE = DEFAULT_DXF.name
MAX_UPLOAD_BYTES = 250 * 1024 * 1024
MAX_JSON_BYTES = 100 * 1024 * 1024
APP_VERSION = "light-editor-2-auto-crs"
GOOGLE_EARTH_CANDIDATES = [
    Path(r"C:\Program Files\Google\Google Earth Pro\client\googleearth.exe"),
    Path(r"C:\Program Files (x86)\Google\Google Earth Pro\client\googleearth.exe"),
]
BLENDER_CANDIDATES = [
    Path(r"D:\Ai 프로그래밍\Blender\blender.exe"),
    Path(r"C:\Program Files\Blender Foundation\Blender\blender.exe"),
]
DEFAULT_SLOPE_BREAKS = [5.0, 10.0, 16.0, 17.0, 25.0, 30.0]
SLOPE_COLORS = ["#48b86f", "#96c75a", "#f2c94c", "#f2994a", "#eb5757", "#9b51e0", "#6d3a75"]
MAX_KML_SLOPE_CELLS = 60000
MAX_KML_SLOPE_FEATURES = 12000
BLENDER_TERRAIN_MAX_GRID = 85
BLENDER_MAX_LINE_POINTS = 120000
BLENDER_MAX_HATCH_PATHS = 6000

PROJECT_CACHE: dict[tuple[str, int | str, float], bytes] = {}
TERRAIN_CACHE: dict[tuple[str, float], "ElevationModel"] = {}
SATELLITE_MANIFEST_CACHE: dict[tuple[str, float, int, int], dict] = {}


@dataclass
class ElevationModel:
    cell_size: float
    bins: dict[tuple[int, int], list[tuple[float, float, float]]]
    min_z: float
    max_z: float
    sample_count: int

    def elevation(self, x: float, y: float) -> float:
        cell_x = math.floor(x / self.cell_size)
        cell_y = math.floor(y / self.cell_size)
        candidates: list[tuple[float, float, float]] = []
        for ring in range(25):
            for offset_x in range(-ring, ring + 1):
                for offset_y in range(-ring, ring + 1):
                    if ring and abs(offset_x) != ring and abs(offset_y) != ring:
                        continue
                    candidates.extend(self.bins.get((cell_x + offset_x, cell_y + offset_y), ()))
            if len(candidates) >= 16:
                break
        if not candidates:
            return 0.0
        nearest = sorted(
            (
                ((sample_x - x) ** 2 + (sample_y - y) ** 2, sample_z)
                for sample_x, sample_y, sample_z in candidates
            ),
            key=lambda item: item[0],
        )[:12]
        if nearest[0][0] < 0.01:
            return nearest[0][1]
        weighted_sum = 0.0
        weight_total = 0.0
        for distance_squared, elevation in nearest:
            weight = 1.0 / (distance_squared + 4.0)
            weighted_sum += elevation * weight
            weight_total += weight
        return weighted_sum / weight_total if weight_total else 0.0


def parse_epsg(query: dict[str, list[str]]) -> int | None:
    raw = query.get("epsg", ["auto"])[0].strip().lower()
    if raw in ("", "auto"):
        return None
    try:
        epsg = int(raw)
    except ValueError:
        return None
    return epsg if epsg in KOREA_CRS else 5187


def project_json(path: Path, epsg: int | None, source_id: str) -> bytes:
    key = (str(path), epsg if epsg is not None else "auto", path.stat().st_mtime)
    if key not in PROJECT_CACHE:
        data = build_project(path, epsg)
        data["sourceId"] = source_id
        data["fileName"] = display_name(source_id)
        PROJECT_CACHE.clear()
        PROJECT_CACHE[key] = json.dumps(data, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    return PROJECT_CACHE[key]


def resolved_epsg(path: Path, requested_epsg: int | None, source_id: str) -> int:
    if requested_epsg is not None:
        return requested_epsg
    return int(json.loads(project_json(path, None, source_id))["epsg"])


def display_name(source_id: str) -> str:
    match = re.match(r"^\d+_(.+)$", source_id)
    return match.group(1) if match else source_id


def sanitize_filename(filename: str) -> str:
    name = Path(filename).name.strip()
    stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", Path(name).stem).strip(" ._") or "project"
    suffix = Path(name).suffix.lower()
    return f"{stem[:80]}{suffix}"


def source_path(source_id: str) -> Path:
    if source_id == DEFAULT_SOURCE:
        return DEFAULT_DXF
    safe_id = Path(source_id).name
    if safe_id != source_id:
        raise ValueError("잘못된 파일 식별자입니다.")
    candidate = (UPLOAD_ROOT / safe_id).resolve()
    if candidate.parent != UPLOAD_ROOT.resolve() or not candidate.is_file():
        raise FileNotFoundError(f"저장된 DXF를 찾을 수 없습니다: {source_id}")
    return candidate


def lon_to_tile_x(lon: float, zoom: int) -> float:
    return (lon + 180.0) / 360.0 * (2**zoom)


def lat_to_tile_y(lat: float, zoom: int) -> float:
    latitude = math.radians(max(-85.05112878, min(85.05112878, lat)))
    return (1.0 - math.asinh(math.tan(latitude)) / math.pi) / 2.0 * (2**zoom)


def tile_x_to_lon(x: float, zoom: int) -> float:
    return x / (2**zoom) * 360.0 - 180.0


def tile_y_to_lat(y: float, zoom: int) -> float:
    return math.degrees(math.atan(math.sinh(math.pi * (1 - 2 * y / (2**zoom)))))


def elevation_model(path: Path, project: dict) -> ElevationModel:
    key = (str(path), path.stat().st_mtime)
    cached = TERRAIN_CACHE.get(key)
    if cached is not None:
        return cached
    samples: list[tuple[float, float, float]] = []
    for feature in project["features"]:
        if feature["kind"] != "polyline" or not feature["layer"].startswith("등고선"):
            continue
        points = feature["points"]
        step = max(1, len(points) // 120)
        for point in points[::step]:
            if abs(point[2]) > 0.001:
                samples.append((point[0], point[1], point[2]))
        if points and abs(points[-1][2]) > 0.001:
            samples.append((points[-1][0], points[-1][1], points[-1][2]))
    if not samples:
        model = ElevationModel(40.0, {}, 0.0, 0.0, 0)
    else:
        cell_size = 40.0
        bins: dict[tuple[int, int], list[tuple[float, float, float]]] = {}
        for sample in samples:
            cell = (math.floor(sample[0] / cell_size), math.floor(sample[1] / cell_size))
            bins.setdefault(cell, []).append(sample)
        model = ElevationModel(
            cell_size=cell_size,
            bins=bins,
            min_z=min(sample[2] for sample in samples),
            max_z=max(sample[2] for sample in samples),
            sample_count=len(samples),
        )
    TERRAIN_CACHE.clear()
    TERRAIN_CACHE[key] = model
    return model


def tile_terrain_mesh(
    tile_x: int,
    tile_y: int,
    zoom: int,
    epsg: int,
    subdivisions: int,
    model: ElevationModel,
) -> dict:
    size = subdivisions + 1
    points = []
    for row in range(size):
        tile_row = tile_y + row / subdivisions
        lat = tile_y_to_lat(tile_row, zoom)
        for column in range(size):
            tile_column = tile_x + column / subdivisions
            lon = tile_x_to_lon(tile_column, zoom)
            x, y = lonlat_to_projected(lon, lat, epsg)
            z = model.elevation(x, y)
            points.append([round(x, 3), round(y, 3), round(z, 3)])
    return {"size": size, "points": points}


def satellite_manifest(path: Path, source_id: str, epsg: int, zoom: int) -> dict:
    cache_key = (str(path), path.stat().st_mtime, epsg, zoom)
    cached = SATELLITE_MANIFEST_CACHE.get(cache_key)
    if cached is not None:
        return cached
    project = json.loads(project_json(path, epsg, source_id))
    model = elevation_model(path, project)
    subdivisions = {15: 5, 16: 7, 17: 10, 18: 12}[zoom]
    geographic = project["bounds"]["wgs84"]
    lon_pad = max((geographic["east"] - geographic["west"]) * 0.12, 0.0002)
    lat_pad = max((geographic["north"] - geographic["south"]) * 0.12, 0.0002)
    west = geographic["west"] - lon_pad
    east = geographic["east"] + lon_pad
    south = geographic["south"] - lat_pad
    north = geographic["north"] + lat_pad
    xmin = max(0, math.floor(lon_to_tile_x(west, zoom)))
    xmax = min(2**zoom - 1, math.floor(lon_to_tile_x(east, zoom)))
    ymin = max(0, math.floor(lat_to_tile_y(north, zoom)))
    ymax = min(2**zoom - 1, math.floor(lat_to_tile_y(south, zoom)))
    tiles = []
    for y in range(ymin, ymax + 1):
        north_lat = tile_y_to_lat(y, zoom)
        south_lat = tile_y_to_lat(y + 1, zoom)
        for x in range(xmin, xmax + 1):
            west_lon = tile_x_to_lon(x, zoom)
            east_lon = tile_x_to_lon(x + 1, zoom)
            nw = lonlat_to_projected(west_lon, north_lat, epsg)
            ne = lonlat_to_projected(east_lon, north_lat, epsg)
            se = lonlat_to_projected(east_lon, south_lat, epsg)
            sw = lonlat_to_projected(west_lon, south_lat, epsg)
            tiles.append(
                {
                    "x": x,
                    "y": y,
                    "url": f"/api/satellite/tile/{zoom}/{x}/{y}.jpeg",
                    "corners": [
                        [round(nw[0], 3), round(nw[1], 3), 0],
                        [round(ne[0], 3), round(ne[1], 3), 0],
                        [round(se[0], 3), round(se[1], 3), 0],
                        [round(sw[0], 3), round(sw[1], 3), 0],
                    ],
                    "terrain": tile_terrain_mesh(x, y, zoom, epsg, subdivisions, model),
                }
            )
    result = {
        "provider": "VWorld",
        "layer": "Satellite",
        "zoom": zoom,
        "tileCount": len(tiles),
        "terrain": {
            "source": "DXF contour interpolation",
            "sampleCount": model.sample_count,
            "minElevation": round(model.min_z, 3),
            "maxElevation": round(model.max_z, 3),
            "subdivisions": subdivisions,
        },
        "tiles": tiles,
    }
    SATELLITE_MANIFEST_CACHE.clear()
    SATELLITE_MANIFEST_CACHE[cache_key] = result
    return result


def fetch_satellite_tile(zoom: int, x: int, y: int) -> bytes:
    cache_path = TILE_CACHE_ROOT / str(zoom) / str(x) / f"{y}.jpeg"
    if cache_path.is_file():
        return cache_path.read_bytes()
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    url = f"https://xdworld.vworld.kr/2d/Satellite/service/{zoom}/{x}/{y}.jpeg"
    request = urllib.request.Request(url, headers={"User-Agent": "LargeProjectAerialViewer/0.2"})
    with urllib.request.urlopen(request, timeout=15) as response:
        body = response.read()
        content_type = response.headers.get_content_type()
    if content_type != "image/jpeg" or len(body) < 500:
        raise ValueError("브이월드 위성 타일 응답이 올바르지 않습니다.")
    cache_path.write_bytes(body)
    return body


def parse_slope_breaks(raw: object) -> list[float]:
    values: list[float] = []
    if isinstance(raw, str):
        parts = re.split(r"[,\s/]+", raw)
    elif isinstance(raw, list):
        parts = raw
    else:
        parts = DEFAULT_SLOPE_BREAKS
    for part in parts:
        try:
            value = float(part)
        except (TypeError, ValueError):
            continue
        if value > 0 and value not in values:
            values.append(value)
    values.sort()
    return values[:10] or DEFAULT_SLOPE_BREAKS[:]


def slope_class_index(degrees: float, breaks: list[float]) -> int:
    for index, upper in enumerate(breaks):
        if degrees <= upper:
            return index
    return len(breaks)


def slope_label(index: int, breaks: list[float]) -> str:
    if index == 0:
        return f"0-{breaks[0]:g}deg"
    if index >= len(breaks):
        return f"{math.floor(breaks[-1]) + 1:g}deg+"
    lower = math.floor(breaks[index - 1]) + 1
    upper = breaks[index]
    return f"{upper:g}deg" if lower == upper else f"{lower:g}-{upper:g}deg"


def is_slope_feature(feature: dict) -> bool:
    layer = str(feature.get("layer", ""))
    return layer.startswith("경사분석") or layer.startswith("Slope ")


def slope_enabled(payload: dict) -> bool:
    raw = payload.get("slope")
    if isinstance(raw, dict):
        return bool(raw.get("enabled", False))
    return True


def collect_slope_samples(project: dict) -> list[tuple[float, float, float]]:
    features = [
        feature for feature in project.get("features", [])
        if feature.get("kind") == "polyline" and isinstance(feature.get("points"), list)
    ]

    def has_finite_z(feature: dict, require_nonzero: bool) -> bool:
        for point in feature.get("points", []):
            if not isinstance(point, list) or len(point) < 3:
                continue
            try:
                z = float(point[2])
            except (TypeError, ValueError):
                continue
            if not require_nonzero or abs(z) > 0.001:
                return True
        return False

    terrain_features = [
        feature for feature in features
        if re.search(r"등고|contour", str(feature.get("layer", "")), re.IGNORECASE)
        and has_finite_z(feature, False)
    ]
    require_nonzero = False
    if not terrain_features:
        terrain_features = [feature for feature in features if has_finite_z(feature, True)]
        require_nonzero = True

    samples: list[tuple[float, float, float]] = []
    for feature in terrain_features:
        points = feature.get("points", [])
        step = max(1, len(points) // 800)
        selected = points[::step]
        if points:
            selected.append(points[-1])
        for point in selected:
            if not isinstance(point, list) or len(point) < 3:
                continue
            try:
                x = float(point[0])
                y = float(point[1])
                z = float(point[2])
            except (TypeError, ValueError):
                continue
            if require_nonzero and abs(z) <= 0.001:
                continue
            samples.append((x, y, z))

    max_samples = 90000
    if len(samples) <= max_samples:
        return samples
    stride = math.ceil(len(samples) / max_samples)
    return [sample for index, sample in enumerate(samples) if index % stride == 0]


def sample_bounds(samples: list[tuple[float, float, float]]) -> tuple[float, float, float, float]:
    xs = [sample[0] for sample in samples]
    ys = [sample[1] for sample in samples]
    return min(xs), min(ys), max(xs), max(ys)


def build_slope_sampler(samples: list[tuple[float, float, float]]):
    bin_size = 40.0
    bins: dict[tuple[int, int], list[tuple[float, float, float]]] = {}
    for sample in samples:
        key = (math.floor(sample[0] / bin_size), math.floor(sample[1] / bin_size))
        bins.setdefault(key, []).append(sample)

    def sample_elevation(x: float, y: float) -> float:
        cell_x = math.floor(x / bin_size)
        cell_y = math.floor(y / bin_size)
        candidates: list[tuple[float, float, float]] = []
        for ring in range(25):
            for offset_x in range(-ring, ring + 1):
                for offset_y in range(-ring, ring + 1):
                    if ring and abs(offset_x) != ring and abs(offset_y) != ring:
                        continue
                    candidates.extend(bins.get((cell_x + offset_x, cell_y + offset_y), ()))
            if len(candidates) >= 16:
                break
        if not candidates:
            return 0.0
        nearest = sorted(
            (((sample_x - x) ** 2 + (sample_y - y) ** 2, sample_z) for sample_x, sample_y, sample_z in candidates),
            key=lambda item: item[0],
        )[:12]
        if nearest[0][0] < 0.01:
            return nearest[0][1]
        weighted_sum = 0.0
        weight_total = 0.0
        for distance_squared, z in nearest:
            weight = 1.0 / (distance_squared + 4.0)
            weighted_sum += z * weight
            weight_total += weight
        return weighted_sum / weight_total if weight_total else 0.0

    return sample_elevation


def slope_feature_from_rect(rect: dict, breaks: list[float], opacity: float) -> dict:
    label = slope_label(int(rect["class_index"]), breaks)
    points = [
        [rect["x0"], rect["y0"], 0.0],
        [rect["x1"], rect["y0"], 0.0],
        [rect["x1"], rect["y1"], 0.0],
        [rect["x0"], rect["y1"], 0.0],
        [rect["x0"], rect["y0"], 0.0],
    ]
    return {
        "kind": "hatch",
        "layer": f"Slope {label}",
        "color": SLOPE_COLORS[int(rect["class_index"]) % len(SLOPE_COLORS)],
        "text": f"Slope {label}",
        "closed": True,
        "pattern": "SOLID",
        "opacity": opacity,
        "points": points,
        "paths": [points],
    }


def build_slope_features(
    bounds: tuple[float, float, float, float],
    sampler,
    breaks: list[float],
    opacity: float,
    cell_size: float,
) -> list[dict]:
    xmin, ymin, xmax, ymax = bounds
    width = max(1.0, xmax - xmin)
    height = max(1.0, ymax - ymin)
    columns = max(1, math.ceil(width / cell_size))
    rows = max(1, math.ceil(height / cell_size))
    node_columns = columns + 1
    node_rows = rows + 1
    elevations = [0.0] * (node_columns * node_rows)
    for row in range(node_rows):
        y = min(ymax, ymin + row * cell_size)
        for column in range(node_columns):
            x = min(xmax, xmin + column * cell_size)
            elevations[row * node_columns + column] = sampler(x, y)

    rectangles: list[dict] = []
    active: dict[tuple[int, float, float], dict] = {}

    def close_run(key: tuple[int, float, float]) -> None:
        run = active.pop(key, None)
        if run is not None:
            rectangles.append(run)

    for row in range(rows):
        y = ymin + row * cell_size
        cell_height = min(cell_size, ymax - y)
        if cell_height <= 0:
            continue
        row_runs: list[dict] = []
        run: dict | None = None

        def flush_row_run() -> None:
            nonlocal run
            if run is not None:
                row_runs.append(run)
                run = None

        for column in range(columns):
            x = xmin + column * cell_size
            cell_width = min(cell_size, xmax - x)
            if cell_width <= 0:
                continue
            top_left = row * node_columns + column
            z00 = elevations[top_left]
            z10 = elevations[top_left + 1]
            z01 = elevations[top_left + node_columns]
            z11 = elevations[top_left + node_columns + 1]
            dzdx = ((z10 + z11) - (z00 + z01)) / (2 * cell_width)
            dzdy = ((z01 + z11) - (z00 + z10)) / (2 * cell_height)
            degrees = math.degrees(math.atan(math.hypot(dzdx, dzdy)))
            class_index = slope_class_index(degrees, breaks)
            if run and run["class_index"] == class_index and abs(run["x1"] - x) < 0.0001:
                run["x1"] = x + cell_width
            else:
                flush_row_run()
                run = {
                    "class_index": class_index,
                    "x0": x,
                    "x1": x + cell_width,
                    "y0": y,
                    "y1": y + cell_height,
                }
        flush_row_run()

        row_keys: set[tuple[int, float, float]] = set()
        for row_run in row_runs:
            key = (int(row_run["class_index"]), round(row_run["x0"], 4), round(row_run["x1"], 4))
            row_keys.add(key)
            existing = active.get(key)
            if existing and abs(existing["y1"] - row_run["y0"]) < 0.0001:
                existing["y1"] = row_run["y1"]
            else:
                close_run(key)
                active[key] = row_run

        for key in list(active):
            if key not in row_keys:
                close_run(key)

    for key in list(active):
        close_run(key)

    return [slope_feature_from_rect(rect, breaks, opacity) for rect in rectangles]


def server_slope_features(payload: dict, epsg: int) -> list[dict]:
    if not slope_enabled(payload):
        return []
    source_id = str(payload.get("sourceId") or ViewerHandler.active_source)
    try:
        path = source_path(source_id)
        project = json.loads(project_json(path, epsg, source_id))
    except Exception:
        return []

    samples = collect_slope_samples(project)
    if len(samples) < 12:
        return []

    slope_config = payload.get("slope") if isinstance(payload.get("slope"), dict) else {}
    breaks = parse_slope_breaks(slope_config.get("breaks", DEFAULT_SLOPE_BREAKS))
    try:
        requested_cell_size = float(slope_config.get("cellSize", 2.0))
    except (TypeError, ValueError):
        requested_cell_size = 2.0
    try:
        opacity = float(slope_config.get("opacity", 0.58))
    except (TypeError, ValueError):
        opacity = 0.58
    opacity = max(0.2, min(0.85, opacity))

    bounds = sample_bounds(samples)
    xmin, ymin, xmax, ymax = bounds
    width = max(1.0, xmax - xmin)
    height = max(1.0, ymax - ymin)
    auto_cell_size = math.ceil(math.sqrt((width * height) / MAX_KML_SLOPE_CELLS))
    cell_size = max(1.0, requested_cell_size, float(auto_cell_size))

    sampler = build_slope_sampler(samples)
    features = build_slope_features(bounds, sampler, breaks, opacity, cell_size)
    for _ in range(4):
        if len(features) <= MAX_KML_SLOPE_FEATURES:
            break
        factor = math.sqrt(len(features) / MAX_KML_SLOPE_FEATURES)
        cell_size = max(cell_size + 1.0, float(math.ceil(cell_size * max(1.2, factor))))
        features = build_slope_features(bounds, sampler, breaks, opacity, cell_size)
    return features


def safe_feature_payload(payload: dict) -> list[dict]:
    features = payload.get("features")
    if not isinstance(features, list):
        raise ValueError("내보낼 도면 객체가 없습니다.")
    result = []
    for feature in features:
        if not isinstance(feature, dict):
            continue
        points = feature.get("points")
        if not isinstance(points, list) or not points:
            continue
        clean_points = []
        for point in points:
            if not isinstance(point, list) or len(point) < 2:
                continue
            try:
                clean_points.append(
                    [
                        float(point[0]),
                        float(point[1]),
                        float(point[2]) if len(point) > 2 else 0.0,
                    ]
                )
            except (TypeError, ValueError):
                continue
        if not clean_points:
            continue
        try:
            opacity = float(feature.get("opacity", 0.4) or 0.4)
        except (TypeError, ValueError):
            opacity = 0.4
        clean_paths = []
        raw_paths = feature.get("paths")
        if isinstance(raw_paths, list):
            for raw_path in raw_paths:
                if not isinstance(raw_path, list):
                    continue
                clean_path = []
                for point in raw_path:
                    if not isinstance(point, list) or len(point) < 2:
                        continue
                    try:
                        clean_path.append(
                            [
                                float(point[0]),
                                float(point[1]),
                                float(point[2]) if len(point) > 2 else 0.0,
                            ]
                        )
                    except (TypeError, ValueError):
                        continue
                if len(clean_path) >= 3:
                    clean_paths.append(clean_path)
        clean_feature = {
            "id": feature.get("id"),
            "kind": str(feature.get("kind", "polyline")),
            "layer": str(feature.get("layer", "0"))[:120] or "0",
            "color": str(feature.get("color", "#dfe6ef")),
            "text": str(feature.get("text", ""))[:500],
            "closed": bool(feature.get("closed", False)),
            "pattern": str(feature.get("pattern", ""))[:120],
            "opacity": max(0.05, min(1.0, opacity)),
            "points": clean_points,
            "paths": clean_paths or None,
        }
        image = feature.get("image")
        if clean_feature["kind"] == "image" and isinstance(image, dict):
            clean_feature["image"] = {
                "path": str(image.get("path", ""))[:1000],
                "fileName": Path(str(image.get("fileName") or image.get("path") or "image")).name[:180],
                "pixelWidth": image.get("pixelWidth"),
                "pixelHeight": image.get("pixelHeight"),
                "available": bool(image.get("available", False)),
                "extension": str(image.get("extension", ""))[:20],
            }
        result.append(clean_feature)
    if not result:
        raise ValueError("내보낼 수 있는 도면 객체가 없습니다.")
    return result


def compact_slope_features_for_kml(features: list[dict]) -> list[dict]:
    result: list[dict] = []
    grouped: dict[tuple[str, str, float], dict] = {}
    for feature in features:
        if not is_slope_feature(feature):
            result.append(feature)
            continue
        opacity = feature["opacity"] if feature["kind"] == "hatch" else 0.4
        key = (feature["layer"], feature["color"], round(opacity, 2))
        group = grouped.get(key)
        if group is None:
            group = {
                "kind": "hatch",
                "layer": feature["layer"],
                "color": feature["color"],
                "text": feature["layer"],
                "closed": True,
                "pattern": "SOLID",
                "opacity": opacity,
                "points": [],
                "paths": [],
            }
            grouped[key] = group
        paths = feature["paths"] or [feature["points"]]
        for path in paths:
            if len(path) >= 3:
                group["paths"].append(path)
    for group in grouped.values():
        if group["paths"]:
            group["points"] = group["paths"][0]
            result.append(group)
    return result


def payload_name(payload: dict, suffix: str) -> str:
    raw = str(payload.get("name", "project"))
    stem = sanitize_filename(Path(raw).stem + suffix)
    return Path(stem).stem


def kml_features_from_payload(payload: dict, epsg: int) -> list[dict]:
    features = safe_feature_payload(payload)
    slope_count = sum(1 for feature in features if is_slope_feature(feature))
    if slope_enabled(payload) and (slope_count == 0 or slope_count > MAX_KML_SLOPE_FEATURES):
        generated_slope = server_slope_features(payload, epsg)
        if generated_slope:
            features = [feature for feature in features if not is_slope_feature(feature)]
            features.extend(generated_slope)
    return compact_slope_features_for_kml(features)


def image_feature_key(feature: dict) -> str:
    image = feature.get("image") if isinstance(feature.get("image"), dict) else {}
    if feature.get("id") is not None:
        return f"id:{feature.get('id')}"
    first = feature.get("points", [[0, 0]])[0]
    return f"{feature.get('layer')}:{image.get('path')}:{first[0]:.3f}:{first[1]:.3f}"


def resolve_payload_image_path(payload: dict, feature: dict) -> Path | None:
    image = feature.get("image") if isinstance(feature.get("image"), dict) else None
    if not image:
        return None
    source_id = str(payload.get("sourceId") or ViewerHandler.active_source)
    try:
        dxf_path = source_path(source_id)
    except Exception:
        dxf_path = ViewerHandler.dxf_path
    resolved = resolve_image_path(str(dxf_path.resolve()), str(image.get("path", "")))
    if not resolved:
        return None
    path = Path(resolved)
    return path if path.is_file() else None


def payload_to_kml(
    payload: dict,
    image_hrefs: dict[str, str] | None = None,
    features: list[dict] | None = None,
) -> str:
    epsg = int(payload.get("epsg", 5187))
    if epsg not in KOREA_CRS:
        epsg = 5187
    features = features if features is not None else kml_features_from_payload(payload, epsg)
    image_hrefs = image_hrefs or {}
    name = html.escape(Path(str(payload.get("name", "project"))).stem)
    style_ids: dict[tuple[str, float, bool], str] = {}
    lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<kml xmlns="http://www.opengis.net/kml/2.2" xmlns:gx="http://www.google.com/kml/ext/2.2">',
        "<Document>",
        f"<name>{name}</name>",
    ]

    def feature_opacity(feature: dict) -> float:
        return feature["opacity"] if feature["kind"] == "hatch" else 0.4

    def feature_style_key(feature: dict) -> tuple[str, float, bool]:
        return (feature["color"], round(feature_opacity(feature), 2), is_slope_feature(feature))

    for feature in features:
        if feature["kind"] == "image":
            continue
        color = feature["color"]
        opacity = feature_opacity(feature)
        slope_style = is_slope_feature(feature)
        style_key = feature_style_key(feature)
        if style_key in style_ids:
            continue
        style_id = f"style_{len(style_ids)}"
        style_ids[style_key] = style_id
        clean = color.lstrip("#")
        if len(clean) != 6:
            clean = "dfe6ef"
        kml_value = f"ff{clean[4:6]}{clean[2:4]}{clean[0:2]}"
        polygon_alpha = f"{round(opacity * 255):02x}"
        polygon_value = f"{polygon_alpha}{clean[4:6]}{clean[2:4]}{clean[0:2]}"
        outline = 0 if slope_style else 1
        line_width = 0.4 if slope_style else 2
        lines.extend(
            [
                f'<Style id="{style_id}">',
                f"<LineStyle><color>{kml_value}</color><width>{line_width:g}</width></LineStyle>",
                f"<PolyStyle><color>{polygon_value}</color><fill>1</fill><outline>{outline}</outline></PolyStyle>",
                f"<IconStyle><color>{kml_value}</color><scale>0.55</scale></IconStyle>",
                "</Style>",
            ]
        )

    center = payload.get("center") or {}
    try:
        center_lon = float(center.get("lon"))
        center_lat = float(center.get("lat"))
        view = payload.get("view") or {}
        heading = -float(view.get("yaw", 0))
        range_value = max(300.0, float(payload.get("range", 1800)))
        lines.extend(
            [
                "<LookAt>",
                f"<longitude>{center_lon:.8f}</longitude>",
                f"<latitude>{center_lat:.8f}</latitude>",
                "<altitude>0</altitude>",
                f"<heading>{heading:.2f}</heading>",
                "<tilt>55</tilt>",
                f"<range>{range_value:.2f}</range>",
                "<altitudeMode>relativeToGround</altitudeMode>",
                "</LookAt>",
            ]
        )
    except (TypeError, ValueError):
        pass

    def append_image_overlay(index: int, feature: dict) -> None:
        href = image_hrefs.get(image_feature_key(feature))
        if not href:
            return
        points = feature["points"][:4]
        if len(points) < 4:
            return
        coords = []
        for point in points:
            lon, lat = projected_to_lonlat(point[0], point[1], epsg)
            coords.append(f"{lon:.8f},{lat:.8f},0")
        image = feature.get("image") if isinstance(feature.get("image"), dict) else {}
        layer = html.escape(feature["layer"])
        feature_name = html.escape(
            str(feature.get("text") or image.get("fileName") or f"Image {index + 1}")
        )
        lines.extend(
            [
                "<GroundOverlay>",
                f"<name>{feature_name}</name>",
                f"<description>Layer: {layer}</description>",
                "<Icon>",
                f"<href>{html.escape(href)}</href>",
                "</Icon>",
                "<gx:LatLonQuad><coordinates>",
                " ".join(coords),
                "</coordinates></gx:LatLonQuad>",
                "</GroundOverlay>",
            ]
        )

    def append_placemark(index: int, feature: dict) -> None:
        if feature["kind"] == "image":
            append_image_overlay(index, feature)
            return
        layer = html.escape(feature["layer"])
        feature_name = html.escape(feature["text"] or f"{feature['layer']} {index + 1}")
        points = feature["points"]
        has_altitude = any(abs(point[2]) > 0.001 for point in points)
        lines.extend(
            [
                "<Placemark>",
                f"<name>{feature_name}</name>",
                f"<description>Layer: {layer}</description>",
                f"<styleUrl>#{style_ids[feature_style_key(feature)]}</styleUrl>",
            ]
        )
        if feature["kind"] in ("point", "text") or len(points) == 1:
            lon, lat = projected_to_lonlat(points[0][0], points[0][1], epsg)
            altitude = max(0.0, points[0][2])
            altitude_mode = "absolute" if has_altitude else "clampToGround"
            lines.append(
                f"<Point><altitudeMode>{altitude_mode}</altitudeMode>"
                f"<coordinates>{lon:.8f},{lat:.8f},{altitude:.2f}</coordinates></Point>"
            )
        elif feature["kind"] == "hatch":
            paths = feature["paths"] or [points]
            altitude_mode = "absolute" if has_altitude else "clampToGround"
            if len(paths) > 1:
                lines.append("<MultiGeometry>")
            for path in paths:
                if len(path) < 3:
                    continue
                ring = list(path)
                if ring[0][:2] != ring[-1][:2]:
                    ring.append(ring[0])
                coords = []
                for point in ring:
                    lon, lat = projected_to_lonlat(point[0], point[1], epsg)
                    coords.append(
                        f"{lon:.8f},{lat:.8f},{max(0.0, point[2]):.2f}"
                    )
                lines.extend(
                    [
                        "<Polygon>",
                        f"<altitudeMode>{altitude_mode}</altitudeMode>",
                        "<tessellate>1</tessellate>",
                        "<outerBoundaryIs><LinearRing><coordinates>",
                        " ".join(coords),
                        "</coordinates></LinearRing></outerBoundaryIs>",
                        "</Polygon>",
                    ]
                )
            if len(paths) > 1:
                lines.append("</MultiGeometry>")
        else:
            coords = []
            for point in points:
                lon, lat = projected_to_lonlat(point[0], point[1], epsg)
                coords.append(f"{lon:.8f},{lat:.8f},{max(0.0, point[2]):.2f}")
            altitude_mode = "absolute" if has_altitude else "clampToGround"
            lines.extend(
                [
                    "<LineString>",
                    f"<altitudeMode>{altitude_mode}</altitudeMode>",
                    "<tessellate>1</tessellate>",
                    "<coordinates>",
                    " ".join(coords),
                    "</coordinates>",
                    "</LineString>",
                ]
            )
        lines.append("</Placemark>")

    drawing_features = [feature for feature in features if not is_slope_feature(feature)]
    slope_features = [feature for feature in features if is_slope_feature(feature)]

    if drawing_features:
        lines.extend(["<Folder>", "<name>Drawing Layers</name>", "<open>1</open>"])
        for index, feature in enumerate(drawing_features):
            append_placemark(index, feature)
        lines.append("</Folder>")

    if slope_features:
        lines.extend(["<Folder>", "<name>Slope Analysis</name>", "<open>1</open>"])
        grouped_slope: dict[str, list[dict]] = {}
        for feature in slope_features:
            grouped_slope.setdefault(feature["layer"], []).append(feature)
        for layer, group in grouped_slope.items():
            lines.extend(
                [
                    "<Folder>",
                    f"<name>{html.escape(layer)}</name>",
                    "<open>0</open>",
                ]
            )
            for index, feature in enumerate(group):
                append_placemark(index, feature)
            lines.append("</Folder>")
        lines.append("</Folder>")
    lines.extend(["</Document>", "</kml>"])
    return "\n".join(lines)


def payload_to_kmz(payload: dict) -> bytes:
    epsg = int(payload.get("epsg", 5187))
    if epsg not in KOREA_CRS:
        epsg = 5187
    features = kml_features_from_payload(payload, epsg)
    image_hrefs: dict[str, str] = {}
    image_files: list[tuple[str, Path]] = []
    for index, feature in enumerate(features):
        if feature["kind"] != "image":
            continue
        image_path = resolve_payload_image_path(payload, feature)
        if image_path is None:
            continue
        suffix = image_path.suffix.lower() or ".jpg"
        arcname = f"files/image_{index}{suffix}"
        image_hrefs[image_feature_key(feature)] = arcname
        image_files.append((arcname, image_path))

    kml = payload_to_kml(payload, image_hrefs=image_hrefs, features=features).encode("utf-8")
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        archive.writestr("doc.kml", kml)
        for arcname, image_path in image_files:
            archive.write(image_path, arcname)
    return output.getvalue()


def dxf_layer_name(value: str) -> str:
    clean = re.sub(r'[<>/\\":;?*|=,\x00-\x1f]', "_", value).strip()
    return clean[:120] or "0"


def payload_to_dxf(payload: dict) -> bytes:
    features = safe_feature_payload(payload)
    lines = [
        "0",
        "SECTION",
        "2",
        "HEADER",
        "9",
        "$ACADVER",
        "1",
        "AC1024",
        "9",
        "$DWGCODEPAGE",
        "3",
        "UTF-8",
        "0",
        "ENDSEC",
        "0",
        "SECTION",
        "2",
        "ENTITIES",
    ]
    for feature in features:
        layer = dxf_layer_name(feature["layer"])
        points = feature["points"]
        if feature["kind"] in ("point", "text") or len(points) == 1:
            point = points[0]
            if feature["kind"] == "text" and feature["text"]:
                lines.extend(
                    [
                        "0", "TEXT", "8", layer,
                        "10", f"{point[0]:.6f}", "20", f"{point[1]:.6f}", "30", f"{point[2]:.6f}",
                        "40", "1.0", "1", feature["text"],
                    ]
                )
            else:
                lines.extend(
                    [
                        "0", "POINT", "8", layer,
                        "10", f"{point[0]:.6f}", "20", f"{point[1]:.6f}", "30", f"{point[2]:.6f}",
                    ]
                )
            continue
        paths = feature["paths"] if feature["kind"] == "hatch" and feature["paths"] else [points]
        for path in paths:
            clean_path = []
            for point in path:
                if isinstance(point, list) and len(point) >= 2:
                    clean_path.append(
                        [
                            float(point[0]),
                            float(point[1]),
                            float(point[2]) if len(point) > 2 else 0.0,
                        ]
                    )
            if len(clean_path) < 2:
                continue
            lines.extend(["0", "POLYLINE", "8", layer, "66", "1", "70", "8"])
            for point in clean_path:
                lines.extend(
                    [
                        "0", "VERTEX", "8", layer,
                        "10", f"{point[0]:.6f}", "20", f"{point[1]:.6f}", "30", f"{point[2]:.6f}",
                        "70", "32",
                    ]
                )
            lines.extend(["0", "SEQEND", "8", layer])
    lines.extend(["0", "ENDSEC", "0", "EOF"])
    return ("\r\n".join(lines) + "\r\n").encode("utf-8")


def rgb_from_hex(hex_color: str) -> tuple[float, float, float]:
    match = re.fullmatch(r"#?([0-9a-fA-F]{6})", str(hex_color or ""))
    value = match.group(1) if match else "dfe6ef"
    return (
        int(value[0:2], 16) / 255.0,
        int(value[2:4], 16) / 255.0,
        int(value[4:6], 16) / 255.0,
    )


def payload_feature_bounds(features: list[dict]) -> tuple[float, float, float, float] | None:
    values: list[tuple[float, float]] = []
    for feature in features:
        for point in feature.get("points") or []:
            if isinstance(point, list) and len(point) >= 2:
                values.append((float(point[0]), float(point[1])))
        for path in feature.get("paths") or []:
            for point in path:
                if isinstance(point, list) and len(point) >= 2:
                    values.append((float(point[0]), float(point[1])))
    if not values:
        return None
    xs = [item[0] for item in values]
    ys = [item[1] for item in values]
    return min(xs), min(ys), max(xs), max(ys)


def blender_origin(payload: dict, features: list[dict], epsg: int) -> tuple[float, float]:
    center = payload.get("center")
    if isinstance(center, dict):
        try:
            lon = float(center.get("lon"))
            lat = float(center.get("lat"))
            return lonlat_to_projected(lon, lat, epsg)
        except (TypeError, ValueError):
            pass
    bounds = payload_feature_bounds(features)
    if bounds:
        xmin, ymin, xmax, ymax = bounds
        return (xmin + xmax) / 2.0, (ymin + ymax) / 2.0
    return 0.0, 0.0


def blender_mesh_bounds(features: list[dict], samples: list[tuple[float, float, float]]) -> tuple[float, float, float, float] | None:
    if samples:
        bounds = sample_bounds(samples)
    else:
        bounds = payload_feature_bounds(features)
    if not bounds:
        return None
    xmin, ymin, xmax, ymax = bounds
    width = max(1.0, xmax - xmin)
    height = max(1.0, ymax - ymin)
    pad = max(width, height) * 0.03
    return xmin - pad, ymin - pad, xmax + pad, ymax + pad


def build_blender_terrain(
    features: list[dict],
    origin_x: float,
    origin_y: float,
    sample_elevation,
    samples: list[tuple[float, float, float]],
) -> dict | None:
    bounds = blender_mesh_bounds(features, samples)
    if bounds is None:
        return None
    xmin, ymin, xmax, ymax = bounds
    width = max(1.0, xmax - xmin)
    height = max(1.0, ymax - ymin)
    longest = max(width, height)
    columns = max(2, min(BLENDER_TERRAIN_MAX_GRID, int(round(width / longest * (BLENDER_TERRAIN_MAX_GRID - 1))) + 1))
    rows = max(2, min(BLENDER_TERRAIN_MAX_GRID, int(round(height / longest * (BLENDER_TERRAIN_MAX_GRID - 1))) + 1))
    vertices = []
    for row in range(rows):
        y = ymin + height * row / (rows - 1)
        for column in range(columns):
            x = xmin + width * column / (columns - 1)
            z = sample_elevation(x, y) if sample_elevation else 0.0
            vertices.append([round(x - origin_x, 3), round(y - origin_y, 3), round(z, 3)])
    faces = []
    for row in range(rows - 1):
        for column in range(columns - 1):
            index = row * columns + column
            faces.append([index, index + 1, index + columns + 1, index + columns])
    return {
        "name": "Existing terrain from DXF contours",
        "vertices": vertices,
        "faces": faces,
        "color": [0.22, 0.34, 0.26, 0.62],
    }


def blender_transform_point(
    point: list[float],
    origin_x: float,
    origin_y: float,
    sample_elevation=None,
    drape_zero_z: bool = False,
    z_offset: float = 0.0,
) -> list[float]:
    x = float(point[0])
    y = float(point[1])
    z = float(point[2]) if len(point) > 2 else 0.0
    if drape_zero_z and sample_elevation and abs(z) <= 0.001:
        z = sample_elevation(x, y)
    return [round(x - origin_x, 3), round(y - origin_y, 3), round(z + z_offset, 3)]


def build_blender_line_groups(
    features: list[dict],
    origin_x: float,
    origin_y: float,
    sample_elevation,
) -> list[dict]:
    line_features = [
        feature for feature in features
        if feature.get("kind") == "polyline" and isinstance(feature.get("points"), list)
    ]
    total_points = sum(len(feature.get("points") or []) for feature in line_features)
    stride = max(1, math.ceil(total_points / BLENDER_MAX_LINE_POINTS)) if total_points else 1
    groups: dict[tuple[str, str], dict] = {}
    for feature in line_features:
        raw_points = feature.get("points") or []
        points = raw_points[::stride]
        if raw_points and points[-1] != raw_points[-1]:
            points.append(raw_points[-1])
        if len(points) < 2:
            continue
        key = (str(feature.get("layer", "0")), str(feature.get("color", "#dfe6ef")))
        group = groups.setdefault(
            key,
            {
                "name": key[0],
                "color": [*rgb_from_hex(key[1]), 1.0],
                "paths": [],
            },
        )
        group["paths"].append(
            [
                blender_transform_point(point, origin_x, origin_y, sample_elevation, True, 0.45)
                for point in points
            ]
        )
    return list(groups.values())


def build_blender_hatch_groups(
    features: list[dict],
    origin_x: float,
    origin_y: float,
    sample_elevation,
) -> list[dict]:
    groups: dict[tuple[str, str, float], dict] = {}
    path_count = 0
    for feature in features:
        if feature.get("kind") != "hatch":
            continue
        raw_paths = feature.get("paths") or [feature.get("points") or []]
        try:
            opacity = float(feature.get("opacity", 0.45) or 0.45)
        except (TypeError, ValueError):
            opacity = 0.45
        key = (str(feature.get("layer", "0")), str(feature.get("color", "#dfe6ef")), round(opacity, 2))
        group = groups.setdefault(
            key,
            {
                "name": key[0],
                "color": [*rgb_from_hex(key[1]), max(0.08, min(1.0, key[2]))],
                "paths": [],
            },
        )
        for raw_path in raw_paths:
            if path_count >= BLENDER_MAX_HATCH_PATHS:
                break
            if not isinstance(raw_path, list) or len(raw_path) < 3:
                continue
            path = [
                blender_transform_point(point, origin_x, origin_y, sample_elevation, True, 0.65)
                for point in raw_path
                if isinstance(point, list) and len(point) >= 2
            ]
            if len(path) >= 3:
                if path[0] == path[-1]:
                    path = path[:-1]
                group["paths"].append(path)
                path_count += 1
        if path_count >= BLENDER_MAX_HATCH_PATHS:
            break
    return [group for group in groups.values() if group["paths"]]


def build_blender_images(
    payload: dict,
    features: list[dict],
    origin_x: float,
    origin_y: float,
    sample_elevation,
) -> list[dict]:
    images = []
    for feature in features:
        if feature.get("kind") != "image":
            continue
        image_path = resolve_payload_image_path(payload, feature)
        if image_path is None:
            continue
        points = feature.get("points") or []
        if len(points) < 4:
            continue
        images.append(
            {
                "name": str(feature.get("text") or image_path.name),
                "path": str(image_path),
                "corners": [
                    blender_transform_point(point, origin_x, origin_y, sample_elevation, True, 0.25)
                    for point in points[:4]
                ],
            }
        )
    return images


def payload_to_blender_script(payload: dict) -> bytes:
    epsg = int(payload.get("epsg", 5187))
    if epsg not in KOREA_CRS:
        epsg = 5187
    features = compact_slope_features_for_kml(safe_feature_payload(payload))
    origin_x, origin_y = blender_origin(payload, features, epsg)
    samples = collect_slope_samples({"features": features})
    sample_elevation = build_slope_sampler(samples) if samples else None
    data = {
        "name": str(payload.get("name", "Aerial Viewer Project")),
        "epsg": epsg,
        "originProjected": [round(origin_x, 3), round(origin_y, 3)],
        "originWgs84": payload.get("center"),
        "terrain": build_blender_terrain(features, origin_x, origin_y, sample_elevation, samples),
        "lineGroups": build_blender_line_groups(features, origin_x, origin_y, sample_elevation),
        "hatchGroups": build_blender_hatch_groups(features, origin_x, origin_y, sample_elevation),
        "images": build_blender_images(payload, features, origin_x, origin_y, sample_elevation),
    }
    script_data = json.dumps(data, ensure_ascii=False)
    script = f'''# -*- coding: utf-8 -*-
"""
Generated by 조감도작성.
Open Blender, go to Scripting, paste/run this file or start Blender with:
blender --python this_file.py
"""
import math
import os
import bpy

DATA = {script_data}


def clear_scene():
    bpy.ops.object.select_all(action="SELECT")
    bpy.ops.object.delete()


def material(name, rgba):
    mat = bpy.data.materials.new(name[:63])
    mat.diffuse_color = rgba
    mat.use_nodes = True
    mat.blend_method = "BLEND" if rgba[3] < 0.99 else "OPAQUE"
    mat.show_transparent_back = True
    bsdf = next((node for node in mat.node_tree.nodes if node.type == "BSDF_PRINCIPLED"), None)
    if bsdf:
        if "Base Color" in bsdf.inputs:
            bsdf.inputs["Base Color"].default_value = rgba
        if "Alpha" in bsdf.inputs:
            bsdf.inputs["Alpha"].default_value = rgba[3]
    return mat


def image_material(name, path):
    mat = bpy.data.materials.new(name[:63])
    mat.diffuse_color = (1, 1, 1, 1)
    mat.use_nodes = True
    mat.blend_method = "BLEND"
    nodes = mat.node_tree.nodes
    links = mat.node_tree.links
    bsdf = next((node for node in nodes if node.type == "BSDF_PRINCIPLED"), None)
    tex = nodes.new(type="ShaderNodeTexImage")
    if os.path.isfile(path):
        tex.image = bpy.data.images.load(path, check_existing=True)
    if bsdf:
        if "Base Color" in bsdf.inputs:
            links.new(tex.outputs["Color"], bsdf.inputs["Base Color"])
        if "Alpha" in bsdf.inputs:
            links.new(tex.outputs["Alpha"], bsdf.inputs["Alpha"])
    return mat


def make_mesh(name, vertices, faces, mat):
    mesh = bpy.data.meshes.new(name[:63] + "Mesh")
    mesh.from_pydata(vertices, [], faces)
    mesh.update()
    obj = bpy.data.objects.new(name[:63], mesh)
    bpy.context.collection.objects.link(obj)
    if mat:
        obj.data.materials.append(mat)
    return obj


def make_curve_group(group):
    curve = bpy.data.curves.new(("CAD " + group["name"])[:63], "CURVE")
    curve.dimensions = "3D"
    curve.resolution_u = 1
    curve.bevel_depth = 0.12
    curve.bevel_resolution = 1
    for path in group["paths"]:
        if len(path) < 2:
            continue
        spline = curve.splines.new("POLY")
        spline.points.add(len(path) - 1)
        for point, coord in zip(spline.points, path):
            point.co = (coord[0], coord[1], coord[2], 1.0)
    obj = bpy.data.objects.new(("CAD " + group["name"])[:63], curve)
    bpy.context.collection.objects.link(obj)
    obj.data.materials.append(material("Line " + group["name"], group["color"]))
    return obj


def make_hatch_group(group):
    vertices = []
    faces = []
    for path in group["paths"]:
        start = len(vertices)
        vertices.extend(path)
        faces.append(list(range(start, start + len(path))))
    return make_mesh("Hatch " + group["name"], vertices, faces, material("Hatch " + group["name"], group["color"]))


def make_image_plane(item):
    obj = make_mesh("Image " + item["name"], item["corners"], [[0, 1, 2, 3]], image_material("Image " + item["name"], item["path"]))
    uv_layer = obj.data.uv_layers.new(name="UVMap")
    uv_values = [(0, 1), (1, 1), (1, 0), (0, 0)]
    for loop, uv in zip(obj.data.polygons[0].loop_indices, uv_values):
        uv_layer.data[loop].uv = uv
    return obj


clear_scene()
terrain = DATA.get("terrain")
if terrain:
    make_mesh(terrain["name"], terrain["vertices"], terrain["faces"], material("Existing Terrain", terrain["color"]))
for image in DATA.get("images", []):
    make_image_plane(image)
for hatch_group in DATA.get("hatchGroups", []):
    make_hatch_group(hatch_group)
for line_group in DATA.get("lineGroups", []):
    make_curve_group(line_group)

diameter = 1000
if terrain and terrain["vertices"]:
    xs = [point[0] for point in terrain["vertices"]]
    ys = [point[1] for point in terrain["vertices"]]
    diameter = max(max(xs) - min(xs), max(ys) - min(ys), 100)
bpy.ops.object.light_add(type="SUN", location=(0, -diameter, diameter))
bpy.context.object.name = "Sun"
bpy.ops.object.camera_add(location=(0, -diameter * 1.25, diameter * 0.85), rotation=(math.radians(58), 0, 0))
bpy.context.scene.camera = bpy.context.object
bpy.context.scene.unit_settings.system = "METRIC"
print("조감도작성 Blender scene loaded:", DATA["name"])
print("Origin EPSG", DATA["epsg"], DATA["originProjected"], DATA.get("originWgs84"))
'''
    return script.encode("utf-8")


def google_earth_executable() -> Path | None:
    for candidate in GOOGLE_EARTH_CANDIDATES:
        if candidate.is_file():
            return candidate
    return None


def blender_executable() -> Path | None:
    for candidate in BLENDER_CANDIDATES:
        if candidate.is_file():
            return candidate
    return None


class ViewerHandler(BaseHTTPRequestHandler):
    dxf_path: Path = DEFAULT_DXF
    active_source: str = DEFAULT_SOURCE

    def log_message(self, format: str, *args) -> None:
        print(f"{self.address_string()} - {format % args}")

    def send_bytes(self, body: bytes, content_type: str, status: int = 200, headers: dict[str, str] | None = None) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        if headers:
            for key, value in headers.items():
                self.send_header(key, value)
        self.end_headers()
        self.wfile.write(body)

    def send_json(self, payload: dict, status: int = 200) -> None:
        self.send_bytes(json.dumps(payload, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8", status)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        query = parse_qs(parsed.query)
        try:
            if parsed.path in ("", "/"):
                return self.serve_static(STATIC_ROOT / "index.html")
            if parsed.path.startswith("/static/"):
                rel = parsed.path.removeprefix("/static/").replace("/", "\\")
                return self.serve_static(STATIC_ROOT / rel)
            if parsed.path == "/api/health":
                return self.send_json(
                    {
                        "ok": True,
                        "version": APP_VERSION,
                        "dxf": str(self.dxf_path),
                        "sourceId": self.active_source,
                    }
                )
            if parsed.path == "/api/project":
                epsg = parse_epsg(query)
                source_id = query.get("source", [self.active_source])[0]
                path = source_path(source_id)
                if query.get("regen", ["0"])[0] == "1":
                    for cache_key in list(PROJECT_CACHE):
                        if cache_key[0] == source_id:
                            PROJECT_CACHE.pop(cache_key, None)
                ViewerHandler.dxf_path = path
                ViewerHandler.active_source = source_id
                return self.send_bytes(project_json(path, epsg, source_id), "application/json; charset=utf-8")
            if parsed.path == "/api/image":
                source_id = query.get("source", [self.active_source])[0]
                raw_id = query.get("id", [""])[0]
                path = source_path(source_id)
                project = json.loads(project_json(path, None, source_id))
                feature = next(
                    (
                        item for item in project.get("features", [])
                        if str(item.get("id")) == str(raw_id) and item.get("kind") == "image"
                    ),
                    None,
                )
                if not feature or not isinstance(feature.get("image"), dict):
                    return self.send_json({"error": "이미지 객체를 찾을 수 없습니다."}, 404)
                image_path = resolve_image_path(str(path.resolve()), str(feature["image"].get("path", "")))
                if not image_path:
                    return self.send_json({"error": "DXF가 참조한 이미지 파일을 찾을 수 없습니다."}, 404)
                resolved = Path(image_path)
                content_type = mimetypes.guess_type(str(resolved))[0] or "application/octet-stream"
                return self.send_bytes(
                    resolved.read_bytes(),
                    content_type,
                    headers={"Cache-Control": "public, max-age=86400"},
                )
            if parsed.path == "/api/export.kml":
                requested_epsg = parse_epsg(query)
                source_id = query.get("source", [self.active_source])[0]
                path = source_path(source_id)
                epsg = resolved_epsg(path, requested_epsg, source_id)
                body = features_to_kml(path, epsg).encode("utf-8")
                return self.send_bytes(
                    body,
                    "application/vnd.google-earth.kml+xml; charset=utf-8",
                    headers={"Content-Disposition": f'attachment; filename="{path.stem}_project.kml"'},
                )
            if parsed.path == "/api/satellite/manifest":
                requested_epsg = parse_epsg(query)
                source_id = query.get("source", [self.active_source])[0]
                path = source_path(source_id)
                epsg = resolved_epsg(path, requested_epsg, source_id)
                zoom = max(15, min(18, int(query.get("zoom", ["17"])[0])))
                return self.send_json(satellite_manifest(path, source_id, epsg, zoom))
            tile_match = re.fullmatch(r"/api/satellite/tile/(\d+)/(\d+)/(\d+)\.jpeg", parsed.path)
            if tile_match:
                zoom, x, y = map(int, tile_match.groups())
                if zoom < 15 or zoom > 18 or x < 0 or y < 0 or x >= 2**zoom or y >= 2**zoom:
                    return self.send_json({"error": "invalid satellite tile"}, 400)
                body = fetch_satellite_tile(zoom, x, y)
                return self.send_bytes(
                    body,
                    "image/jpeg",
                    headers={"Cache-Control": "public, max-age=2592000"},
                )
            return self.send_json({"error": "not found"}, 404)
        except Exception as exc:
            return self.send_json({"error": str(exc)}, 500)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        query = parse_qs(parsed.query)
        try:
            if parsed.path in ("/api/export/dxf", "/api/export/blender", "/api/blender/open", "/api/export/kmz", "/api/google-earth/open"):
                length = int(self.headers.get("Content-Length", "0"))
                if length <= 0 or length > MAX_JSON_BYTES:
                    return self.send_json({"error": "내보내기 데이터 크기가 올바르지 않습니다."}, 413)
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                base_name = sanitize_filename(Path(str(payload.get("name", "project"))).stem) or "project"
                if parsed.path == "/api/export/dxf":
                    body = payload_to_dxf(payload)
                    return self.send_bytes(
                        body,
                        "application/dxf",
                        headers={
                            "Content-Disposition": 'attachment; filename="edited_project.dxf"'
                        },
                    )
                if parsed.path == "/api/export/blender":
                    body = payload_to_blender_script(payload)
                    return self.send_bytes(
                        body,
                        "text/x-python; charset=utf-8",
                        headers={
                            "Content-Disposition": f'attachment; filename="{base_name}_blender.py"'
                        },
                    )
                if parsed.path == "/api/blender/open":
                    executable = blender_executable()
                    if executable is None:
                        return self.send_json(
                            {
                                "error": "Blender를 찾을 수 없습니다. D:\\Ai 프로그래밍\\Blender\\blender.exe 설치를 확인해 주세요.",
                                "scriptAvailable": True,
                            },
                            404,
                        )
                    EXPORT_ROOT.mkdir(parents=True, exist_ok=True)
                    output_path = EXPORT_ROOT / f"{base_name}_blender_{int(time.time())}.py"
                    output_path.write_bytes(payload_to_blender_script(payload))
                    subprocess.Popen(
                        [str(executable), "--python", str(output_path)],
                        cwd=str(executable.parent),
                        close_fds=True,
                    )
                    return self.send_json(
                        {
                            "ok": True,
                            "path": str(output_path),
                            "blender": str(executable),
                            "message": "Blender에서 현재 도면을 열었습니다.",
                        }
                    )
                kmz = payload_to_kmz(payload)
                if parsed.path == "/api/export/kmz":
                    return self.send_bytes(
                        kmz,
                        "application/vnd.google-earth.kmz",
                        headers={
                            "Content-Disposition": 'attachment; filename="google_earth_project.kmz"'
                        },
                    )
                executable = google_earth_executable()
                if executable is None:
                    return self.send_json(
                        {
                            "error": "Google Earth Pro를 찾을 수 없습니다.",
                            "kmzAvailable": True,
                        },
                        404,
                    )
                EXPORT_ROOT.mkdir(parents=True, exist_ok=True)
                output_path = EXPORT_ROOT / f"{base_name}_{int(time.time())}.kmz"
                output_path.write_bytes(kmz)
                subprocess.Popen(
                    [str(executable), str(output_path)],
                    cwd=str(executable.parent),
                    close_fds=True,
                )
                return self.send_json(
                    {
                        "ok": True,
                        "path": str(output_path),
                        "message": "Google Earth Pro에서 현재 도면을 열었습니다.",
                    }
                )
            if parsed.path != "/api/upload":
                return self.send_json({"error": "not found"}, 404)
            raw_name = unquote(self.headers.get("X-File-Name", "project.dxf"))
            filename = sanitize_filename(raw_name)
            suffix = Path(filename).suffix.lower()
            if suffix == ".dwg":
                return self.send_json(
                    {"error": "DWG는 변환 엔진 연동 전입니다. 현재는 DXF 파일을 불러와 주세요."},
                    415,
                )
            if suffix != ".dxf":
                return self.send_json({"error": "DXF 파일만 불러올 수 있습니다."}, 415)
            length = int(self.headers.get("Content-Length", "0"))
            if length <= 0 or length > MAX_UPLOAD_BYTES:
                return self.send_json({"error": "파일 크기가 올바르지 않거나 250MB를 초과했습니다."}, 413)
            body = self.rfile.read(length)
            if len(body) != length:
                return self.send_json({"error": "파일 전송이 완료되지 않았습니다."}, 400)
            UPLOAD_ROOT.mkdir(parents=True, exist_ok=True)
            source_id = f"{int(time.time())}_{filename}"
            path = UPLOAD_ROOT / source_id
            path.write_bytes(body)
            epsg = parse_epsg(query)
            try:
                payload = project_json(path, epsg, source_id)
            except Exception:
                path.unlink(missing_ok=True)
                raise
            ViewerHandler.dxf_path = path
            ViewerHandler.active_source = source_id
            return self.send_bytes(payload, "application/json; charset=utf-8")
        except Exception as exc:
            return self.send_json({"error": str(exc)}, 500)

    def serve_static(self, path: Path) -> None:
        resolved = path.resolve()
        if not str(resolved).startswith(str(STATIC_ROOT.resolve())) or not resolved.exists() or not resolved.is_file():
            return self.send_json({"error": "static file not found"}, 404)
        content_type = mimetypes.guess_type(str(resolved))[0] or "application/octet-stream"
        if resolved.suffix == ".js":
            content_type = "text/javascript; charset=utf-8"
        elif resolved.suffix in (".html", ".css"):
            content_type = f"text/{resolved.suffix[1:]}; charset=utf-8"
        self.send_bytes(resolved.read_bytes(), content_type)


def main() -> None:
    parser = argparse.ArgumentParser(description="Large project aerial view prototype")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--dxf", type=Path, default=DEFAULT_DXF)
    args = parser.parse_args()

    if not args.dxf.exists():
        raise SystemExit(f"DXF file not found: {args.dxf}")
    ViewerHandler.dxf_path = args.dxf.resolve()
    ViewerHandler.active_source = DEFAULT_SOURCE if args.dxf.resolve() == DEFAULT_DXF.resolve() else args.dxf.name
    UPLOAD_ROOT.mkdir(parents=True, exist_ok=True)
    TILE_CACHE_ROOT.mkdir(parents=True, exist_ok=True)
    server = ThreadingHTTPServer((args.host, args.port), ViewerHandler)
    print(f"Viewer running at http://{args.host}:{args.port}")
    print(f"DXF: {ViewerHandler.dxf_path}")
    server.serve_forever()


if __name__ == "__main__":
    main()
