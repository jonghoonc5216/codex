from __future__ import annotations

import html
import math
import re
from collections import Counter
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Iterable


A_GRS80 = 6378137.0
F_GRS80 = 1 / 298.257222101
E2_GRS80 = F_GRS80 * (2 - F_GRS80)
EP2_GRS80 = E2_GRS80 / (1 - E2_GRS80)


KOREA_CRS = {
    5179: {
        "name": "Korea 2000 / Unified CS",
        "lat0": 38.0,
        "lon0": 127.5,
        "k0": 0.9996,
        "x0": 1_000_000.0,
        "y0": 2_000_000.0,
    },
    5185: {
        "name": "Korea 2000 / West Belt 2010",
        "lat0": 38.0,
        "lon0": 125.0,
        "k0": 1.0,
        "x0": 200_000.0,
        "y0": 600_000.0,
    },
    5186: {
        "name": "Korea 2000 / Central Belt 2010",
        "lat0": 38.0,
        "lon0": 127.0,
        "k0": 1.0,
        "x0": 200_000.0,
        "y0": 600_000.0,
    },
    5187: {
        "name": "Korea 2000 / East Belt 2010",
        "lat0": 38.0,
        "lon0": 129.0,
        "k0": 1.0,
        "x0": 200_000.0,
        "y0": 600_000.0,
    },
    5188: {
        "name": "Korea 2000 / East Sea Belt 2010",
        "lat0": 38.0,
        "lon0": 131.0,
        "k0": 1.0,
        "x0": 200_000.0,
        "y0": 600_000.0,
    },
}


ACI_COLORS = {
    1: "#f0514a",
    2: "#f5cf4b",
    3: "#3fbf62",
    4: "#45c7d8",
    5: "#4f7cff",
    6: "#d35bd6",
    7: "#f4f5f7",
    8: "#8f99a8",
    9: "#c3cad4",
    10: "#e06b5c",
    30: "#ef9f43",
    90: "#60b86d",
    130: "#43b6b2",
    170: "#4b8fe3",
    210: "#8f70d6",
    250: "#d7dce4",
}


@dataclass
class Feature:
    kind: str
    layer: str
    points: list[list[float]]
    color: str = "#dfe6ef"
    text: str | None = None
    closed: bool = False
    paths: list[list[list[float]]] | None = None
    pattern: str | None = None
    image: dict | None = None


@dataclass
class Block:
    name: str
    base: list[float]
    entities: list[tuple[str, list[tuple[str, str]]]]


@dataclass
class ImageDef:
    handle: str
    path: str
    width: float | None = None
    height: float | None = None


@dataclass(frozen=True)
class Transform:
    a: float = 1.0
    b: float = 0.0
    c: float = 0.0
    d: float = 1.0
    tx: float = 0.0
    ty: float = 0.0
    sz: float = 1.0
    tz: float = 0.0

    def point(self, point: list[float]) -> list[float]:
        return [
            self.a * point[0] + self.c * point[1] + self.tx,
            self.b * point[0] + self.d * point[1] + self.ty,
            self.sz * point[2] + self.tz,
        ]

    def then(self, child: "Transform") -> "Transform":
        return Transform(
            a=self.a * child.a + self.c * child.b,
            b=self.b * child.a + self.d * child.b,
            c=self.a * child.c + self.c * child.d,
            d=self.b * child.c + self.d * child.d,
            tx=self.a * child.tx + self.c * child.ty + self.tx,
            ty=self.b * child.tx + self.d * child.ty + self.ty,
            sz=self.sz * child.sz,
            tz=self.sz * child.tz + self.tz,
        )


@lru_cache(maxsize=16)
def detect_dxf_encoding(path_string: str, mtime: float) -> str:
    data = Path(path_string).read_bytes()
    try:
        data.decode("utf-8")
        return "utf-8"
    except UnicodeDecodeError:
        return "cp949"


def dxf_pairs(path: Path) -> Iterable[tuple[str, str]]:
    encoding = detect_dxf_encoding(str(path.resolve()), path.stat().st_mtime)
    with path.open("r", encoding=encoding, errors="replace", newline="") as handle:
        while True:
            code = handle.readline()
            if not code:
                return
            value = handle.readline()
            if not value:
                return
            yield code.strip(), value.rstrip("\r\n")


def safe_float(value: str | None, default: float | None = None) -> float | None:
    if value is None:
        return default
    try:
        return float(value.strip())
    except (TypeError, ValueError):
        return default


def safe_int(value: str | None, default: int | None = None) -> int | None:
    if value is None:
        return default
    try:
        return int(float(value.strip()))
    except (TypeError, ValueError):
        return default


def first_tag(tags: list[tuple[str, str]], code: str, default: str | None = None) -> str | None:
    for tag_code, value in tags:
        if tag_code == code:
            return value
    return default


def color_from_tags(tags: list[tuple[str, str]]) -> str:
    color_index = safe_int(first_tag(tags, "62"), None)
    if color_index is None or color_index == 256:
        return "#dfe6ef"
    return ACI_COLORS.get(abs(color_index), "#dfe6ef")


def layer_from_tags(tags: list[tuple[str, str]]) -> str:
    return first_tag(tags, "8", "0") or "0"


def handle_from_tags(tags: list[tuple[str, str]]) -> str | None:
    return first_tag(tags, "5")


def get_xyz(tags: list[tuple[str, str]], x_code: str = "10", y_code: str = "20", z_code: str = "30") -> list[float] | None:
    x = safe_float(first_tag(tags, x_code))
    y = safe_float(first_tag(tags, y_code))
    z = safe_float(first_tag(tags, z_code), 0.0)
    if x is None or y is None:
        return None
    return [x, y, z or 0.0]


def collect_repeated_points(tags: list[tuple[str, str]], x_code: str = "10", y_code: str = "20", z_code: str = "30") -> list[list[float]]:
    points: list[list[float]] = []
    pending_x: float | None = None
    pending_y: float | None = None
    pending_z: float = 0.0
    for code, value in tags:
        if code == x_code:
            if pending_x is not None and pending_y is not None:
                points.append([pending_x, pending_y, pending_z])
            pending_x = safe_float(value)
            pending_y = None
            pending_z = 0.0
        elif code == y_code and pending_x is not None:
            pending_y = safe_float(value)
        elif code == z_code and pending_x is not None:
            pending_z = safe_float(value, 0.0) or 0.0
    if pending_x is not None and pending_y is not None:
        points.append([pending_x, pending_y, pending_z])
    return [point for point in points if point[0] is not None and point[1] is not None]


def lwpolyline_points(
    tags: list[tuple[str, str]],
    closed: bool = False,
) -> list[list[float]]:
    elevation = safe_float(first_tag(tags, "38"), 0.0) or 0.0
    vertices: list[tuple[list[float], float]] = []
    current: list[float | None] | None = None
    bulge = 0.0
    for code, value in tags:
        if code == "10":
            if current and current[0] is not None and current[1] is not None:
                vertices.append(
                    (
                        [float(current[0]), float(current[1]), float(current[2] or 0.0)],
                        bulge,
                    )
                )
            current = [safe_float(value), None, elevation]
            bulge = 0.0
        elif current is not None and code == "20":
            current[1] = safe_float(value)
        elif current is not None and code == "30":
            current[2] = (safe_float(value, 0.0) or 0.0) + elevation
        elif current is not None and code == "42":
            bulge = safe_float(value, 0.0) or 0.0
    if current and current[0] is not None and current[1] is not None:
        vertices.append(
            (
                [float(current[0]), float(current[1]), float(current[2] or 0.0)],
                bulge,
            )
        )
    if len(vertices) < 2:
        return [point for point, _bulge in vertices]

    points: list[list[float]] = []
    segment_count = len(vertices) if closed else len(vertices) - 1
    for index in range(segment_count):
        start, segment_bulge = vertices[index]
        end = vertices[(index + 1) % len(vertices)][0]
        segment = bulge_segment(start, end, segment_bulge)
        points.extend(segment if not points else segment[1:])
    return points


def arc_points(
    center: list[float],
    radius: float,
    start_angle: float,
    end_angle: float,
    segments: int = 36,
) -> list[list[float]]:
    if end_angle < start_angle:
        end_angle += 360.0
    sweep = max(1.0, end_angle - start_angle)
    count = max(4, min(96, int(segments * sweep / 360.0) + 2))
    points = []
    for index in range(count):
        t = index / (count - 1)
        angle = math.radians(start_angle + sweep * t)
        points.append(
            [
                center[0] + radius * math.cos(angle),
                center[1] + radius * math.sin(angle),
                center[2],
            ]
        )
    return points


def circle_points(center: list[float], radius: float, segments: int = 72) -> list[list[float]]:
    return [
        [
            center[0] + radius * math.cos(2 * math.pi * index / segments),
            center[1] + radius * math.sin(2 * math.pi * index / segments),
            center[2],
        ]
        for index in range(segments + 1)
    ]


def hatch_arc_points(
    center: list[float],
    radius: float,
    start_angle: float,
    end_angle: float,
    ccw: bool,
) -> list[list[float]]:
    sweep = (end_angle - start_angle) % 360.0
    if sweep < 1e-7:
        return circle_points(center, radius)
    count = max(4, min(96, int(36 * sweep / 360.0) + 2))
    y_direction = 1.0 if ccw else -1.0
    points = []
    for index in range(count):
        angle = math.radians(start_angle + sweep * index / (count - 1))
        points.append(
            [
                center[0] + radius * math.cos(angle),
                center[1] + y_direction * radius * math.sin(angle),
                center[2],
            ]
        )
    return points


def ellipse_points(tags: list[tuple[str, str]], segments: int = 72) -> list[list[float]]:
    center = get_xyz(tags)
    major = get_xyz(tags, "11", "21", "31")
    ratio = safe_float(first_tag(tags, "40"), 1.0) or 1.0
    start_param = safe_float(first_tag(tags, "41"), 0.0) or 0.0
    end_param = safe_float(first_tag(tags, "42"), 2 * math.pi) or 2 * math.pi
    if center is None or major is None:
        return []
    if end_param < start_param:
        end_param += 2 * math.pi
    mx, my, mz = major
    minor = [-my * ratio, mx * ratio, mz]
    count = max(8, min(128, int(segments * (end_param - start_param) / (2 * math.pi)) + 2))
    points = []
    for index in range(count):
        t = index / (count - 1)
        angle = start_param + (end_param - start_param) * t
        points.append(
            [
                center[0] + mx * math.cos(angle) + minor[0] * math.sin(angle),
                center[1] + my * math.cos(angle) + minor[1] * math.sin(angle),
                center[2],
            ]
        )
    return points


def read_header_metadata(path: Path) -> dict:
    metadata: dict = {"extmin": None, "extmax": None, "insunits": None, "codepage": None}
    section: str | None = None
    current_var: str | None = None
    ext_values: dict[str, dict[str, float]] = {"$EXTMIN": {}, "$EXTMAX": {}}
    pairs = iter(dxf_pairs(path))
    for code, value in pairs:
        if code == "0" and value == "SECTION":
            try:
                next_code, next_value = next(pairs)
            except StopIteration:
                break
            section = next_value if next_code == "2" else None
            continue
        if code == "0" and value == "ENDSEC":
            if section == "HEADER":
                break
            section = None
            continue
        if section != "HEADER":
            continue
        if code == "9":
            current_var = value
            continue
        if current_var in ("$EXTMIN", "$EXTMAX") and code in ("10", "20", "30"):
            axis = {"10": "x", "20": "y", "30": "z"}[code]
            val = safe_float(value)
            if val is not None:
                ext_values[current_var][axis] = val
        elif current_var == "$INSUNITS" and code == "70":
            metadata["insunits"] = safe_int(value)
        elif current_var == "$DWGCODEPAGE" and code == "3":
            metadata["codepage"] = value
    if "x" in ext_values["$EXTMIN"] and "y" in ext_values["$EXTMIN"]:
        metadata["extmin"] = ext_values["$EXTMIN"]
    if "x" in ext_values["$EXTMAX"] and "y" in ext_values["$EXTMAX"]:
        metadata["extmax"] = ext_values["$EXTMAX"]
    return metadata


def section_entity_stream(path: Path, target_section: str) -> Iterable[tuple[str, list[tuple[str, str]]]]:
    section: str | None = None
    current_type: str | None = None
    current_tags: list[tuple[str, str]] = []
    pairs = iter(dxf_pairs(path))
    for code, value in pairs:
        if code == "0" and value == "SECTION":
            if current_type is not None and section == target_section:
                yield current_type, current_tags
            current_type = None
            current_tags = []
            try:
                next_code, next_value = next(pairs)
            except StopIteration:
                return
            section = next_value if next_code == "2" else None
            continue
        if code == "0" and value == "ENDSEC":
            if current_type is not None and section == target_section:
                yield current_type, current_tags
            current_type = None
            current_tags = []
            if section == target_section:
                return
            section = None
            continue
        if section != target_section:
            continue
        if code == "0":
            if current_type is not None:
                yield current_type, current_tags
            current_type = value
            current_tags = []
        else:
            current_tags.append((code, value))
    if current_type is not None and section == target_section:
        yield current_type, current_tags


def entity_stream(path: Path) -> Iterable[tuple[str, list[tuple[str, str]]]]:
    return section_entity_stream(path, "ENTITIES")


def read_blocks(path: Path) -> dict[str, Block]:
    blocks: dict[str, Block] = {}
    current: Block | None = None
    for entity_type, tags in section_entity_stream(path, "BLOCKS"):
        if entity_type == "BLOCK":
            name = first_tag(tags, "2") or first_tag(tags, "3")
            if not name:
                current = None
                continue
            current = Block(name=name, base=get_xyz(tags) or [0.0, 0.0, 0.0], entities=[])
            blocks[name] = current
            continue
        if entity_type == "ENDBLK":
            current = None
            continue
        if current is not None:
            current.entities.append((entity_type, tags))
    return blocks


def feature_bounds(points: list[list[float]]) -> tuple[float, float, float, float] | None:
    if not points:
        return None
    return (
        min(point[0] for point in points),
        min(point[1] for point in points),
        max(point[0] for point in points),
        max(point[1] for point in points),
    )


def valid_header_extents(extmin: dict | None, extmax: dict | None) -> bool:
    if not extmin or not extmax:
        return False
    values = [extmin.get("x"), extmin.get("y"), extmax.get("x"), extmax.get("y")]
    if any(value is None or not math.isfinite(value) or abs(value) >= 1e15 for value in values):
        return False
    return extmin["x"] < extmax["x"] and extmin["y"] < extmax["y"]


def valid_z_extents(extmin: dict | None, extmax: dict | None) -> bool:
    if not extmin or not extmax or "z" not in extmin or "z" not in extmax:
        return False
    zmin = extmin["z"]
    zmax = extmax["z"]
    return (
        math.isfinite(zmin)
        and math.isfinite(zmax)
        and abs(zmin) < 1e15
        and abs(zmax) < 1e15
        and zmin <= zmax
    )


def dominant_feature_bounds(features: list[Feature]) -> tuple[float, float, float, float] | None:
    points = [point for feature in features for point in feature.points]
    if not points:
        return None

    def axis_bounds(axis: int) -> tuple[float, float]:
        values = sorted(point[axis] for point in points if math.isfinite(point[axis]))
        if not values:
            return 0.0, 0.0
        if len(values) < 100:
            return values[0], values[-1]

        cluster = values
        for _attempt in range(3):
            if len(cluster) < 100:
                break
            gaps = [
                (cluster[index + 1] - cluster[index], index)
                for index in range(len(cluster) - 1)
            ]
            largest_gap, split_index = max(gaps)
            left = cluster[: split_index + 1]
            right = cluster[split_index + 1 :]
            dominant = left if len(left) >= len(right) else right
            dominant_ratio = len(dominant) / len(cluster)
            dominant_span = max(1.0, dominant[-1] - dominant[0])
            if dominant_ratio < 0.8 or largest_gap <= max(1_000.0, dominant_span * 2):
                break
            cluster = dominant

        low = cluster[0]
        high = cluster[-1]
        pad = max(20.0, (high - low) * 0.01)
        return low - pad, high + pad

    xmin, xmax = axis_bounds(0)
    ymin, ymax = axis_bounds(1)
    return xmin, ymin, xmax, ymax


def intersects_bounds(points: list[list[float]], bounds: tuple[float, float, float, float] | None) -> bool:
    if bounds is None:
        return True
    own = feature_bounds(points)
    if own is None:
        return False
    xmin, ymin, xmax, ymax = own
    bxmin, bymin, bxmax, bymax = bounds
    return xmax >= bxmin and xmin <= bxmax and ymax >= bymin and ymin <= bymax


def bulge_segment(start: list[float], end: list[float], bulge: float) -> list[list[float]]:
    if abs(bulge) < 1e-9:
        return [start, end]
    dx = end[0] - start[0]
    dy = end[1] - start[1]
    chord = math.hypot(dx, dy)
    if chord < 1e-9:
        return [start]
    midpoint_x = (start[0] + end[0]) / 2
    midpoint_y = (start[1] + end[1]) / 2
    center_offset = chord * (1 - bulge * bulge) / (4 * bulge)
    center_x = midpoint_x - dy / chord * center_offset
    center_y = midpoint_y + dx / chord * center_offset
    start_angle = math.atan2(start[1] - center_y, start[0] - center_x)
    sweep = 4 * math.atan(bulge)
    count = max(4, min(256, int(abs(sweep) / (math.pi / 60)) + 2))
    return [
        [
            center_x + math.hypot(start[0] - center_x, start[1] - center_y) * math.cos(start_angle + sweep * i / (count - 1)),
            center_y + math.hypot(start[0] - center_x, start[1] - center_y) * math.sin(start_angle + sweep * i / (count - 1)),
            start[2] + (end[2] - start[2]) * i / (count - 1),
        ]
        for i in range(count)
    ]


def hatch_paths(tags: list[tuple[str, str]]) -> list[list[list[float]]]:
    try:
        index = next(i for i, (code, _value) in enumerate(tags) if code == "91") + 1
    except StopIteration:
        return []
    loop_count = safe_int(tags[index - 1][1], 0) or 0
    paths: list[list[list[float]]] = []
    for _loop in range(loop_count):
        while index < len(tags) and tags[index][0] != "92":
            index += 1
        if index >= len(tags):
            break
        flags = safe_int(tags[index][1], 0) or 0
        index += 1
        path: list[list[float]] = []
        if flags & 2:
            while index < len(tags) and tags[index][0] != "93":
                index += 1
            if index >= len(tags):
                break
            vertex_count = safe_int(tags[index][1], 0) or 0
            index += 1
            vertices: list[tuple[list[float], float]] = []
            for _vertex in range(vertex_count):
                while index < len(tags) and tags[index][0] != "10":
                    index += 1
                if index >= len(tags):
                    break
                x = safe_float(tags[index][1])
                index += 1
                y: float | None = None
                bulge = 0.0
                while index < len(tags) and tags[index][0] not in ("10", "97", "92"):
                    code, value = tags[index]
                    if code == "20":
                        y = safe_float(value)
                    elif code == "42":
                        bulge = safe_float(value, 0.0) or 0.0
                    index += 1
                if x is not None and y is not None:
                    vertices.append(([x, y, 0.0], bulge))
            if len(vertices) >= 2:
                for vertex_index, (start, bulge) in enumerate(vertices):
                    end = vertices[(vertex_index + 1) % len(vertices)][0]
                    segment = bulge_segment(start, end, bulge)
                    path.extend(segment if not path else segment[1:])
        else:
            while index < len(tags) and tags[index][0] != "93":
                index += 1
            if index >= len(tags):
                break
            edge_count = safe_int(tags[index][1], 0) or 0
            index += 1
            for _edge in range(edge_count):
                while index < len(tags) and tags[index][0] != "72":
                    index += 1
                if index >= len(tags):
                    break
                edge_type = safe_int(tags[index][1], 0) or 0
                index += 1
                edge_tags: list[tuple[str, str]] = []
                while index < len(tags) and tags[index][0] not in ("72", "97", "92"):
                    edge_tags.append(tags[index])
                    index += 1
                edge_points: list[list[float]] = []
                if edge_type == 1:
                    start = get_xyz(edge_tags, "10", "20", "30")
                    end = get_xyz(edge_tags, "11", "21", "31")
                    if start and end:
                        edge_points = [start, end]
                elif edge_type == 2:
                    center = get_xyz(edge_tags)
                    radius = safe_float(first_tag(edge_tags, "40"))
                    start_angle = safe_float(first_tag(edge_tags, "50"))
                    end_angle = safe_float(first_tag(edge_tags, "51"))
                    ccw = bool(safe_int(first_tag(edge_tags, "73"), 1))
                    if center and radius and start_angle is not None and end_angle is not None:
                        edge_points = hatch_arc_points(
                            center,
                            radius,
                            start_angle,
                            end_angle,
                            ccw,
                        )
                elif edge_type == 3:
                    edge_points = ellipse_points(edge_tags)
                elif edge_type == 4:
                    fit_points = collect_repeated_points(edge_tags, "11", "21", "31")
                    control_points = collect_repeated_points(edge_tags, "10", "20", "30")
                    edge_points = fit_points if len(fit_points) >= 2 else control_points
                if edge_points:
                    if path and path[-1][:2] == edge_points[0][:2]:
                        path.extend(edge_points[1:])
                    else:
                        path.extend(edge_points)
        while index < len(tags) and tags[index][0] not in ("97", "92"):
            index += 1
        if index < len(tags) and tags[index][0] == "97":
            index += 1
        if len(path) >= 3:
            if path[0][:2] != path[-1][:2]:
                path.append(path[0][:])
            paths.append(path)
    return paths


def insert_transform(tags: list[tuple[str, str]], block: Block) -> Transform:
    insertion = get_xyz(tags) or [0.0, 0.0, 0.0]
    sx = safe_float(first_tag(tags, "41"), 1.0) or 1.0
    sy = safe_float(first_tag(tags, "42"), 1.0) or 1.0
    sz = safe_float(first_tag(tags, "43"), 1.0) or 1.0
    angle = math.radians(safe_float(first_tag(tags, "50"), 0.0) or 0.0)
    cos_angle = math.cos(angle)
    sin_angle = math.sin(angle)
    a = cos_angle * sx
    b = sin_angle * sx
    c = -sin_angle * sy
    d = cos_angle * sy
    return Transform(
        a=a,
        b=b,
        c=c,
        d=d,
        tx=insertion[0] - a * block.base[0] - c * block.base[1],
        ty=insertion[1] - b * block.base[0] - d * block.base[1],
        sz=sz,
        tz=insertion[2] - sz * block.base[2],
    )


def transformed_feature(feature: Feature, transform: Transform, inherited_layer: str | None = None) -> Feature:
    layer = inherited_layer if feature.layer == "0" and inherited_layer else feature.layer
    paths = None
    if feature.paths:
        paths = [[transform.point(point) for point in path] for path in feature.paths]
        points = [point for path in paths for point in path]
    else:
        points = [transform.point(point) for point in feature.points]
    return Feature(
        kind=feature.kind,
        layer=layer,
        points=points,
        color=feature.color,
        text=feature.text,
        closed=feature.closed,
        paths=paths,
        pattern=feature.pattern,
        image=dict(feature.image) if feature.image else None,
    )


def parse_image_entity(tags: list[tuple[str, str]], image_defs: dict[str, ImageDef]) -> Feature | None:
    insertion = get_xyz(tags, "10", "20", "30")
    u_vector = get_xyz(tags, "11", "21", "31")
    v_vector = get_xyz(tags, "12", "22", "32")
    pixel_width = safe_float(first_tag(tags, "13"))
    pixel_height = safe_float(first_tag(tags, "23"))
    if not insertion or not u_vector or not v_vector or not pixel_width or not pixel_height:
        return None

    image_def_handle = first_tag(tags, "340", "") or ""
    image_def = image_defs.get(image_def_handle)
    x0, y0, z0 = insertion
    ux, uy, uz = u_vector
    vx, vy, vz = v_vector
    u = [ux * pixel_width, uy * pixel_width, uz * pixel_width]
    v = [vx * pixel_height, vy * pixel_height, vz * pixel_height]
    p0 = [x0, y0, z0]
    p1 = [x0 + u[0], y0 + u[1], z0 + u[2]]
    p2 = [x0 + u[0] + v[0], y0 + u[1] + v[1], z0 + u[2] + v[2]]
    p3 = [x0 + v[0], y0 + v[1], z0 + v[2]]
    raw_path = image_def.path if image_def else ""
    return Feature(
        "image",
        layer_from_tags(tags),
        [p0, p1, p2, p3, p0[:]],
        color_from_tags(tags),
        text=Path(raw_path).name if raw_path else "DXF image",
        closed=True,
        image={
            "handle": handle_from_tags(tags) or "",
            "definitionHandle": image_def_handle,
            "path": raw_path,
            "fileName": Path(raw_path).name if raw_path else "image",
            "pixelWidth": pixel_width,
            "pixelHeight": pixel_height,
        },
    )


def parse_entity(entity_type: str, tags: list[tuple[str, str]], image_defs: dict[str, ImageDef] | None = None) -> Feature | None:
    layer = layer_from_tags(tags)
    color = color_from_tags(tags)
    if entity_type == "IMAGE":
        return parse_image_entity(tags, image_defs or {})
    if entity_type == "LINE":
        start = get_xyz(tags, "10", "20", "30")
        end = get_xyz(tags, "11", "21", "31")
        if start and end:
            return Feature("polyline", layer, [start, end], color)
    if entity_type == "LWPOLYLINE":
        closed = bool((safe_int(first_tag(tags, "70"), 0) or 0) & 1)
        points = lwpolyline_points(tags, closed)
        if len(points) >= 2:
            if closed and points[0] != points[-1]:
                points.append(points[0][:])
            return Feature("polyline", layer, points, color, closed=closed)
    if entity_type == "ARC":
        center = get_xyz(tags)
        radius = safe_float(first_tag(tags, "40"))
        start = safe_float(first_tag(tags, "50"))
        end = safe_float(first_tag(tags, "51"))
        if center and radius and start is not None and end is not None:
            return Feature("polyline", layer, arc_points(center, radius, start, end), color)
    if entity_type == "CIRCLE":
        center = get_xyz(tags)
        radius = safe_float(first_tag(tags, "40"))
        if center and radius:
            return Feature("polyline", layer, circle_points(center, radius), color, closed=True)
    if entity_type == "ELLIPSE":
        points = ellipse_points(tags)
        if len(points) >= 2:
            return Feature("polyline", layer, points, color)
    if entity_type == "SPLINE":
        fit_points = collect_repeated_points(tags, "11", "21", "31")
        control_points = collect_repeated_points(tags, "10", "20", "30")
        points = fit_points if len(fit_points) >= 2 else control_points
        if len(points) >= 2:
            return Feature("polyline", layer, points, color)
    if entity_type == "SOLID":
        points = []
        for x_code, y_code, z_code in (("10", "20", "30"), ("11", "21", "31"), ("12", "22", "32"), ("13", "23", "33")):
            point = get_xyz(tags, x_code, y_code, z_code)
            if point:
                points.append(point)
        if len(points) >= 3:
            points.append(points[0][:])
            return Feature("polyline", layer, points, color, closed=True)
    if entity_type == "HATCH":
        paths = hatch_paths(tags)
        if paths:
            return Feature(
                "hatch",
                layer,
                [point for path in paths for point in path],
                color,
                closed=True,
                paths=paths,
                pattern=first_tag(tags, "2", "SOLID") or "SOLID",
            )
    if entity_type in ("POINT", "INSERT"):
        point = get_xyz(tags)
        if point:
            return Feature("point", layer, [point], color)
    if entity_type in ("TEXT", "MTEXT"):
        point = get_xyz(tags)
        value = first_tag(tags, "1") or first_tag(tags, "3")
        if point and value:
            clean_text = value.replace("\\P", " ").replace("\\~", " ").strip()
            return Feature("text", layer, [point], color, clean_text[:80])
    return None


def expand_entities(
    entities: list[tuple[str, list[tuple[str, str]]]],
    blocks: dict[str, Block],
    image_defs: dict[str, ImageDef] | None = None,
    transform: Transform | None = None,
    inherited_layer: str | None = None,
    depth: int = 0,
) -> tuple[list[Feature], dict[str, int]]:
    active_transform = transform or Transform()
    features: list[Feature] = []
    ignored: dict[str, int] = {}
    open_polyline: Feature | None = None

    def add_ignored(name: str) -> None:
        ignored[name] = ignored.get(name, 0) + 1

    for entity_type, tags in entities:
        if entity_type == "POLYLINE":
            if open_polyline and len(open_polyline.points) >= 2:
                features.append(transformed_feature(open_polyline, active_transform, inherited_layer))
            open_polyline = Feature("polyline", layer_from_tags(tags), [], color_from_tags(tags))
            continue
        if entity_type == "VERTEX" and open_polyline is not None:
            point = get_xyz(tags)
            if point:
                open_polyline.points.append(point)
            continue
        if entity_type == "SEQEND":
            if open_polyline and len(open_polyline.points) >= 2:
                features.append(transformed_feature(open_polyline, active_transform, inherited_layer))
            open_polyline = None
            continue
        if entity_type == "INSERT":
            block_name = first_tag(tags, "2")
            block = blocks.get(block_name or "")
            if block is not None and depth < 8:
                insert_layer = layer_from_tags(tags)
                child_layer = inherited_layer if insert_layer == "0" and inherited_layer else insert_layer
                child_transform = active_transform.then(insert_transform(tags, block))
                child_features, child_ignored = expand_entities(
                    block.entities,
                    blocks,
                    image_defs,
                    child_transform,
                    child_layer,
                    depth + 1,
                )
                features.extend(child_features)
                for key, count in child_ignored.items():
                    ignored[key] = ignored.get(key, 0) + count
                continue
            add_ignored("INSERT_missing_block" if block is None else "INSERT_depth_limit")
            continue

        feature = parse_entity(entity_type, tags, image_defs)
        if feature is None:
            add_ignored(entity_type)
            continue
        features.append(transformed_feature(feature, active_transform, inherited_layer))

    if open_polyline and len(open_polyline.points) >= 2:
        features.append(transformed_feature(open_polyline, active_transform, inherited_layer))
    return features, ignored


def read_image_defs(path: Path) -> dict[str, ImageDef]:
    image_defs: dict[str, ImageDef] = {}
    for entity_type, tags in section_entity_stream(path, "OBJECTS"):
        if entity_type != "IMAGEDEF":
            continue
        handle = handle_from_tags(tags)
        raw_path = first_tag(tags, "1")
        if not handle or not raw_path:
            continue
        image_defs[handle] = ImageDef(
            handle=handle,
            path=raw_path,
            width=safe_float(first_tag(tags, "10")),
            height=safe_float(first_tag(tags, "20")),
        )
    return image_defs


IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tif", ".tiff"}


@lru_cache(maxsize=512)
def resolve_image_path(dxf_path_string: str, image_ref: str) -> str | None:
    dxf_path = Path(dxf_path_string)
    raw = str(image_ref or "").strip().strip('"')
    if not raw:
        return None
    normalized = raw.replace("\\", "/")
    image_path = Path(raw)
    candidates: list[Path] = []
    if image_path.is_absolute():
        candidates.append(image_path)
    else:
        candidates.append(dxf_path.parent / raw)
        candidates.append(dxf_path.parent / normalized)
        name = Path(normalized).name
        search_roots = [
            dxf_path.parent,
            Path.cwd(),
            Path.home() / "Desktop" / "Ai 실습" / "★대형프로젝트 조감도",
        ]
        for root in search_roots:
            if not root.exists():
                continue
            candidates.append(root / raw)
            candidates.append(root / normalized)
            candidates.append(root / name)
        for root in search_roots:
            if not root.exists():
                continue
            try:
                match = next(root.rglob(name), None)
            except OSError:
                match = None
            if match is not None:
                candidates.append(match)

    for candidate in candidates:
        try:
            resolved = candidate.resolve()
        except OSError:
            continue
        if resolved.is_file() and resolved.suffix.lower() in IMAGE_SUFFIXES:
            return str(resolved)
    return None


HAPBON1_MARKER = "\ud569\ubcf81"
KNOWN_AERIAL_SAMPLE_ROOT = (
    Path.home()
    / "Desktop"
    / "Ai \uc2e4\uc2b5"
    / "\u2605\ub300\ud615\ud504\ub85c\uc81d\ud2b8 \uc870\uac10\ub3c4"
)


def upload_original_name(path: Path) -> str:
    match = re.match(r"^\d+_(.+)$", path.name)
    return match.group(1) if match else path.name


def companion_gbk_candidates(path: Path) -> list[Path]:
    names = [path.name, upload_original_name(path)]
    stems: list[str] = []
    for name in names:
        stem = Path(name).stem
        if stem and stem not in stems:
            stems.append(stem)

    roots: list[Path] = []
    for root in (path.parent, KNOWN_AERIAL_SAMPLE_ROOT):
        if root.exists() and root not in roots:
            roots.append(root)

    candidates: list[Path] = []
    for root in roots:
        for stem in stems:
            candidate = root / f"{stem}.gbk"
            if candidate.exists() and candidate not in candidates:
                candidates.append(candidate)
        try:
            gbk_files = list(root.glob("*.gbk"))
        except OSError:
            gbk_files = []
        for gbk_file in gbk_files:
            if gbk_file in candidates:
                continue
            if any(gbk_file.stem == stem or gbk_file.stem.endswith(stem) for stem in stems):
                candidates.append(gbk_file)
    return candidates


def image_resolved_path(path: Path, feature: Feature) -> Path | None:
    if feature.kind != "image" or not feature.image:
        return None
    resolved = resolve_image_path(str(path.resolve()), str(feature.image.get("path", "")))
    return Path(resolved) if resolved else None


def image_name_matches_hapbon1(resolved: Path | None, feature: Feature) -> bool:
    names = []
    if resolved is not None:
        names.append(resolved.name)
    if feature.image:
        names.append(str(feature.image.get("fileName", "")))
        names.append(str(feature.image.get("path", "")))
    return any(HAPBON1_MARKER in name for name in names)


def normalize_image_feature(feature: Feature, resolved: Path) -> Feature:
    if feature.image is None:
        return feature
    feature.image = {
        **feature.image,
        "path": str(resolved),
        "fileName": resolved.name,
    }
    feature.text = resolved.name
    return feature


def filter_available_image_features(path: Path, features: list[Feature], ignored: dict[str, int]) -> list[Feature]:
    filtered: list[Feature] = []
    missing = 0
    for feature in features:
        if feature.kind != "image":
            filtered.append(feature)
            continue
        resolved = image_resolved_path(path, feature)
        if resolved is None:
            missing += 1
            continue
        filtered.append(normalize_image_feature(feature, resolved))
    if missing:
        ignored["image_missing_file"] = ignored.get("image_missing_file", 0) + missing
    return filtered


def supplement_hapbon1_image_from_gbk(path: Path, features: list[Feature], ignored: dict[str, int]) -> list[Feature]:
    for feature in features:
        resolved = image_resolved_path(path, feature)
        if resolved is not None and image_name_matches_hapbon1(resolved, feature):
            return features

    for companion in companion_gbk_candidates(path):
        if companion.resolve() == path.resolve():
            continue
        try:
            blocks = read_blocks(companion)
            image_defs = read_image_defs(companion)
            companion_features, _companion_ignored = expand_entities(
                list(entity_stream(companion)),
                blocks,
                image_defs,
            )
        except Exception:
            continue

        available: list[tuple[Feature, Path]] = []
        hapbon1: list[tuple[Feature, Path]] = []
        for feature in companion_features:
            resolved = image_resolved_path(companion, feature)
            if resolved is None:
                continue
            available.append((feature, resolved))
            if image_name_matches_hapbon1(resolved, feature):
                hapbon1.append((feature, resolved))

        selected = hapbon1 or (available if len(available) == 1 else [])
        if not selected:
            continue

        additions = [normalize_image_feature(feature, resolved) for feature, resolved in selected]
        ignored["image_supplemented_from_gbk"] = ignored.get("image_supplemented_from_gbk", 0) + len(additions)
        return [*features, *additions]

    return features


def parse_dxf(path: Path) -> tuple[dict, list[Feature], dict]:
    metadata = read_header_metadata(path)
    extmin = metadata.get("extmin")
    extmax = metadata.get("extmax")
    padded_bounds: tuple[float, float, float, float] | None = None
    header_bounds_valid = valid_header_extents(extmin, extmax)
    if header_bounds_valid:
        width = max(1.0, extmax["x"] - extmin["x"])
        height = max(1.0, extmax["y"] - extmin["y"])
        pad = max(50.0, max(width, height) * 0.08)
        padded_bounds = (
            extmin["x"] - pad,
            extmin["y"] - pad,
            extmax["x"] + pad,
            extmax["y"] + pad,
        )

    blocks = read_blocks(path)
    image_defs = read_image_defs(path)
    features, ignored = expand_entities(list(entity_stream(path)), blocks, image_defs)
    if not header_bounds_valid:
        padded_bounds = dominant_feature_bounds(features)
        ignored["invalid_header_extents"] = 1
    if valid_z_extents(extmin, extmax):
        z_pad = max(20.0, (extmax["z"] - extmin["z"]) * 0.25)
        z_low = extmin["z"] - z_pad
        z_high = extmax["z"] + z_pad
        flattened = 0
        for feature in features:
            for point in feature.points:
                if point[2] < z_low or point[2] > z_high:
                    point[2] = 0.0
                    flattened += 1
        if flattened:
            ignored["z_outliers_flattened"] = flattened
    features = supplement_hapbon1_image_from_gbk(path, features, ignored)
    features = filter_available_image_features(path, features, ignored)
    filtered: list[Feature] = []
    for feature in features:
        if intersects_bounds(feature.points, padded_bounds):
            filtered.append(feature)
        else:
            key = f"{feature.kind}_out_of_bounds"
            ignored[key] = ignored.get(key, 0) + 1
    return metadata, filtered, ignored


def meridional_arc(phi: float) -> float:
    e2 = E2_GRS80
    e4 = e2 * e2
    e6 = e4 * e2
    return A_GRS80 * (
        (1 - e2 / 4 - 3 * e4 / 64 - 5 * e6 / 256) * phi
        - (3 * e2 / 8 + 3 * e4 / 32 + 45 * e6 / 1024) * math.sin(2 * phi)
        + (15 * e4 / 256 + 45 * e6 / 1024) * math.sin(4 * phi)
        - (35 * e6 / 3072) * math.sin(6 * phi)
    )


def projected_to_lonlat(x: float, y: float, epsg: int = 5187) -> tuple[float, float]:
    if epsg not in KOREA_CRS:
        raise ValueError(f"Unsupported EPSG:{epsg}")
    params = KOREA_CRS[epsg]
    phi0 = math.radians(params["lat0"])
    lambda0 = math.radians(params["lon0"])
    k0 = params["k0"]
    x0 = params["x0"]
    y0 = params["y0"]
    m0 = meridional_arc(phi0)
    m = m0 + (y - y0) / k0
    mu = m / (A_GRS80 * (1 - E2_GRS80 / 4 - 3 * E2_GRS80**2 / 64 - 5 * E2_GRS80**3 / 256))
    e1 = (1 - math.sqrt(1 - E2_GRS80)) / (1 + math.sqrt(1 - E2_GRS80))
    fp = (
        mu
        + (3 * e1 / 2 - 27 * e1**3 / 32) * math.sin(2 * mu)
        + (21 * e1**2 / 16 - 55 * e1**4 / 32) * math.sin(4 * mu)
        + (151 * e1**3 / 96) * math.sin(6 * mu)
        + (1097 * e1**4 / 512) * math.sin(8 * mu)
    )
    c1 = EP2_GRS80 * math.cos(fp) ** 2
    t1 = math.tan(fp) ** 2
    n1 = A_GRS80 / math.sqrt(1 - E2_GRS80 * math.sin(fp) ** 2)
    r1 = A_GRS80 * (1 - E2_GRS80) / (1 - E2_GRS80 * math.sin(fp) ** 2) ** 1.5
    d = (x - x0) / (n1 * k0)
    lat = fp - (n1 * math.tan(fp) / r1) * (
        d**2 / 2
        - (5 + 3 * t1 + 10 * c1 - 4 * c1**2 - 9 * EP2_GRS80) * d**4 / 24
        + (61 + 90 * t1 + 298 * c1 + 45 * t1**2 - 252 * EP2_GRS80 - 3 * c1**2) * d**6 / 720
    )
    lon = lambda0 + (
        d
        - (1 + 2 * t1 + c1) * d**3 / 6
        + (5 - 2 * c1 + 28 * t1 - 3 * c1**2 + 8 * EP2_GRS80 + 24 * t1**2) * d**5 / 120
    ) / math.cos(fp)
    return math.degrees(lon), math.degrees(lat)


def lonlat_to_projected(lon: float, lat: float, epsg: int = 5187) -> tuple[float, float]:
    if epsg not in KOREA_CRS:
        raise ValueError(f"Unsupported EPSG:{epsg}")
    params = KOREA_CRS[epsg]
    phi = math.radians(lat)
    lambda_value = math.radians(lon)
    phi0 = math.radians(params["lat0"])
    lambda0 = math.radians(params["lon0"])
    k0 = params["k0"]
    sin_phi = math.sin(phi)
    cos_phi = math.cos(phi)
    tan_phi = math.tan(phi)
    n = A_GRS80 / math.sqrt(1 - E2_GRS80 * sin_phi * sin_phi)
    t = tan_phi * tan_phi
    c = EP2_GRS80 * cos_phi * cos_phi
    a_term = (lambda_value - lambda0) * cos_phi
    m = meridional_arc(phi)
    m0 = meridional_arc(phi0)
    x = params["x0"] + k0 * n * (
        a_term
        + (1 - t + c) * a_term**3 / 6
        + (5 - 18 * t + t * t + 72 * c - 58 * EP2_GRS80) * a_term**5 / 120
    )
    y = params["y0"] + k0 * (
        m
        - m0
        + n
        * tan_phi
        * (
            a_term**2 / 2
            + (5 - t + 9 * c + 4 * c * c) * a_term**4 / 24
            + (61 - 58 * t + t * t + 600 * c - 330 * EP2_GRS80) * a_term**6 / 720
        )
    )
    return x, y


def round_point(point: list[float]) -> list[float]:
    return [round(point[0], 3), round(point[1], 3), round(point[2], 3)]


def detect_korea_crs(path: Path, features: list[Feature]) -> tuple[int, dict]:
    hints: Counter[int] = Counter()
    hint_sources: dict[int, str] = {}
    keyword_hints = {
        5179: ("통합원점", "unified cs"),
        5185: ("서부원점", "west belt"),
        5186: ("중부원점", "central belt"),
        5187: ("동부원점", "east belt"),
        5188: ("동해원점", "east sea belt"),
    }
    strings = [path.name]
    for feature in features:
        strings.append(feature.layer)
        if feature.text:
            strings.append(feature.text)
    for value in strings:
        lowered = value.lower()
        for epsg in KOREA_CRS:
            if re.search(rf"(?<!\d){epsg}(?!\d)", lowered):
                hints[epsg] += 1
                hint_sources.setdefault(epsg, value)
            for keyword in keyword_hints[epsg]:
                if keyword in lowered:
                    hints[epsg] += 1
                    hint_sources.setdefault(epsg, value)

    if hints:
        epsg, count = hints.most_common(1)[0]
        tied = [candidate for candidate, candidate_count in hints.items() if candidate_count == count]
        if len(tied) == 1:
            return epsg, {
                "mode": "auto",
                "confidence": "high",
                "reason": f"도면 표기 '{hint_sources[epsg]}'에서 EPSG:{epsg}을 확인했습니다.",
            }

    points = [point for feature in features for point in feature.points]
    xs = sorted(point[0] for point in points if math.isfinite(point[0]))
    ys = sorted(point[1] for point in points if math.isfinite(point[1]))
    if not xs or not ys:
        return 5187, {
            "mode": "auto",
            "confidence": "fallback",
            "reason": "좌표 표본이 부족하여 부서 기본값 EPSG:5187을 적용했습니다.",
        }
    center_x = xs[len(xs) // 2]
    center_y = ys[len(ys) // 2]

    if 500_000 <= center_x <= 1_500_000 and 1_000_000 <= center_y <= 3_000_000:
        return 5179, {
            "mode": "auto",
            "confidence": "high",
            "reason": "통합원점 좌표 규모(X 약 100만, Y 약 200만)를 감지했습니다.",
        }

    plausible: list[int] = []
    for candidate in (5185, 5186, 5187, 5188):
        lon, lat = projected_to_lonlat(center_x, center_y, candidate)
        if 32.0 <= lat <= 44.5 and 123.0 <= lon <= 133.5:
            plausible.append(candidate)
    if len(plausible) == 1:
        epsg = plausible[0]
        return epsg, {
            "mode": "auto",
            "confidence": "medium",
            "reason": f"좌표 중심이 한반도 범위에 들어오는 원점 EPSG:{epsg}을 적용했습니다.",
        }

    return 5187, {
        "mode": "auto",
        "confidence": "fallback",
        "reason": "원점 구분 표기가 없어 부서 기본값 EPSG:5187(동부원점)을 적용했습니다.",
    }


def build_project(path: Path, epsg: int | None = None) -> dict:
    metadata, features, ignored = parse_dxf(path)
    if not features:
        raise ValueError("No supported DXF entities were found.")

    if epsg is None:
        epsg, crs_detection = detect_korea_crs(path, features)
    else:
        crs_detection = {
            "mode": "manual",
            "confidence": "manual",
            "reason": f"사용자가 EPSG:{epsg}을 선택했습니다.",
        }

    all_points = [point for feature in features for point in feature.points]
    xmin = min(point[0] for point in all_points)
    ymin = min(point[1] for point in all_points)
    zmin = min(point[2] for point in all_points)
    xmax = max(point[0] for point in all_points)
    ymax = max(point[1] for point in all_points)
    zmax = max(point[2] for point in all_points)
    center_x = (xmin + xmax) / 2
    center_y = (ymin + ymax) / 2
    center_lon, center_lat = projected_to_lonlat(center_x, center_y, epsg)
    sw_lon, sw_lat = projected_to_lonlat(xmin, ymin, epsg)
    ne_lon, ne_lat = projected_to_lonlat(xmax, ymax, epsg)

    layers: dict[str, dict] = {}
    for feature in features:
        layer = layers.setdefault(
            feature.layer,
            {"name": feature.layer, "count": 0, "points": 0, "color": feature.color},
        )
        layer["count"] += 1
        layer["points"] += len(feature.points)
        if layer["color"] == "#dfe6ef" and feature.color != "#dfe6ef":
            layer["color"] = feature.color

    serialized = []
    for index, feature in enumerate(features):
        item = {
            "id": index,
            "kind": feature.kind,
            "layer": feature.layer,
            "color": feature.color,
            "points": [round_point(point) for point in feature.points],
            "closed": feature.closed,
        }
        if feature.text:
            item["text"] = feature.text
        if feature.paths:
            item["paths"] = [
                [round_point(point) for point in path]
                for path in feature.paths
            ]
        if feature.pattern:
            item["pattern"] = feature.pattern
        if feature.kind == "hatch":
            item["opacity"] = 0.24 if (feature.pattern or "").upper() == "SOLID" else 0.55
        if feature.kind == "image" and feature.image:
            resolved_image = resolve_image_path(str(path.resolve()), str(feature.image.get("path", "")))
            item["image"] = {
                **feature.image,
                "available": bool(resolved_image),
                "extension": Path(resolved_image or feature.image.get("path", "")).suffix.lower(),
            }
            item["opacity"] = 1.0
        serialized.append(item)

    return {
        "source": str(path),
        "fileName": path.name,
        "epsg": epsg,
        "crsName": KOREA_CRS[epsg]["name"],
        "crsDetection": crs_detection,
        "crsOptions": [{"epsg": key, "name": value["name"]} for key, value in KOREA_CRS.items()],
        "header": metadata,
        "bounds": {
            "projected": {
                "xmin": round(xmin, 3),
                "ymin": round(ymin, 3),
                "zmin": round(zmin, 3),
                "xmax": round(xmax, 3),
                "ymax": round(ymax, 3),
                "zmax": round(zmax, 3),
            },
            "wgs84": {
                "west": round(min(sw_lon, ne_lon), 8),
                "south": round(min(sw_lat, ne_lat), 8),
                "east": round(max(sw_lon, ne_lon), 8),
                "north": round(max(sw_lat, ne_lat), 8),
            },
        },
        "center": {
            "projected": {"x": round(center_x, 3), "y": round(center_y, 3)},
            "wgs84": {"lon": round(center_lon, 8), "lat": round(center_lat, 8)},
        },
        "stats": {
            "features": len(features),
            "points": len(all_points),
            "layers": len(layers),
            "ignored": ignored,
        },
        "layers": sorted(layers.values(), key=lambda layer: (-layer["count"], layer["name"])),
        "features": serialized,
    }


def kml_color(hex_color: str, alpha: str = "ff") -> str:
    clean = hex_color.lstrip("#")
    if len(clean) != 6:
        clean = "dfe6ef"
    rr, gg, bb = clean[0:2], clean[2:4], clean[4:6]
    return f"{alpha}{bb}{gg}{rr}"


def features_to_kml(path: Path, epsg: int = 5187) -> str:
    metadata, features, _ignored = parse_dxf(path)
    name = html.escape(path.stem)
    lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<kml xmlns="http://www.opengis.net/kml/2.2">',
        "<Document>",
        f"<name>{name}</name>",
    ]
    style_ids: dict[str, str] = {}
    for feature in features:
        if feature.color not in style_ids:
            style_id = f"style_{len(style_ids)}"
            style_ids[feature.color] = style_id
            lines.extend(
                [
                    f'<Style id="{style_id}">',
                    f"<LineStyle><color>{kml_color(feature.color)}</color><width>2</width></LineStyle>",
                    f"<PolyStyle><color>{kml_color(feature.color, '66')}</color><fill>1</fill><outline>1</outline></PolyStyle>",
                    f"<IconStyle><color>{kml_color(feature.color)}</color><scale>0.6</scale></IconStyle>",
                    "</Style>",
                ]
            )

    for index, feature in enumerate(features):
        layer = html.escape(feature.layer)
        placemark_name = html.escape(feature.text or f"{layer} {index + 1}")
        lines.append("<Placemark>")
        lines.append(f"<name>{placemark_name}</name>")
        lines.append(f"<description>Layer: {layer}</description>")
        lines.append(f"<styleUrl>#{style_ids[feature.color]}</styleUrl>")
        if feature.kind in ("point", "text"):
            lon, lat = projected_to_lonlat(feature.points[0][0], feature.points[0][1], epsg)
            lines.append(f"<Point><coordinates>{lon:.8f},{lat:.8f},0</coordinates></Point>")
        elif feature.kind == "hatch":
            paths = feature.paths or [feature.points]
            if len(paths) > 1:
                lines.append("<MultiGeometry>")
            for path_points in paths:
                if len(path_points) < 3:
                    continue
                ring = list(path_points)
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
                        "<Polygon><tessellate>1</tessellate>",
                        "<outerBoundaryIs><LinearRing><coordinates>",
                        " ".join(coords),
                        "</coordinates></LinearRing></outerBoundaryIs></Polygon>",
                    ]
                )
            if len(paths) > 1:
                lines.append("</MultiGeometry>")
        else:
            coords = []
            for point in feature.points:
                lon, lat = projected_to_lonlat(point[0], point[1], epsg)
                coords.append(f"{lon:.8f},{lat:.8f},{max(0.0, point[2]):.2f}")
            lines.append("<LineString><tessellate>1</tessellate>")
            lines.append("<coordinates>")
            lines.append(" ".join(coords))
            lines.append("</coordinates></LineString>")
        lines.append("</Placemark>")
    if metadata.get("extmin") and metadata.get("extmax"):
        extmin = metadata["extmin"]
        extmax = metadata["extmax"]
        west, south = projected_to_lonlat(extmin["x"], extmin["y"], epsg)
        east, north = projected_to_lonlat(extmax["x"], extmax["y"], epsg)
        lines.extend(
            [
                "<Placemark>",
                "<name>DXF header extent</name>",
                "<Style><LineStyle><color>ff00ffff</color><width>3</width></LineStyle></Style>",
                "<LineString><tessellate>1</tessellate><coordinates>",
                f"{west:.8f},{south:.8f},0 {east:.8f},{south:.8f},0 {east:.8f},{north:.8f},0 {west:.8f},{north:.8f},0 {west:.8f},{south:.8f},0",
                "</coordinates></LineString>",
                "</Placemark>",
            ]
        )
    lines.extend(["</Document>", "</kml>"])
    return "\n".join(lines)
