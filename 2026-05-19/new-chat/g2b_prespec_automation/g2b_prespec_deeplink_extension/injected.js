(function () {
  if (window.__g2bPrespecDetailHelperLoaded) {
    return;
  }
  window.__g2bPrespecDetailHelperLoaded = true;
  window.__g2bPrespecDetailHelperState = {
    opened: false,
    ready: false,
    error: "",
    request: null
  };

  function getRequest() {
    const params = new URLSearchParams(window.location.search || "");
    const specNo = params.get("bfSpecRgstNo") || params.get("g2bPrespecRegNo") || params.get("bfSpecRegNo");
    if (!specNo) {
      return null;
    }
    return {
      bfSpecRegNo: specNo,
      bfSpecRgstNo: specNo,
      prcmBsneSeCd: params.get("g2bPrespecBsneSeCd") || params.get("prcmBsneSeCd") || "",
      bsnsDivNm: params.get("bsnsDivNm") || ""
    };
  }

  function isReady() {
    try {
      const ready = Boolean(
        window.com &&
          typeof window.com.gfnOpenMenu === "function" &&
          window._gcm &&
          window._gcm.Global &&
          typeof window._gcm.Global.getMainLoaded === "function" &&
          window._gcm.Global.getMainLoaded()
      );
      window.__g2bPrespecDetailHelperState.ready = ready;
      return ready;
    } catch (error) {
      window.__g2bPrespecDetailHelperState.error = String(error);
      return false;
    }
  }

  let opened = false;
  function openDetail() {
    const request = getRequest();
    if (!request || opened || !isReady()) {
      return;
    }
    window.__g2bPrespecDetailHelperState.request = request;

    const param = {
      bfSpecRegNo: request.bfSpecRegNo,
      bfSpecRgstNo: request.bfSpecRgstNo
    };
    if (request.prcmBsneSeCd) {
      param.prcmBsneSeCd = request.prcmBsneSeCd;
    }
    if (request.bsnsDivNm) {
      param.bsnsDivNm = request.bsnsDivNm;
    }

    opened = true;
    window.__g2bPrespecDetailHelperState.opened = true;
    try {
      window.com.gfnOpenMenu("PRVA004_02", {
        isHistory: true,
        param
      });
    } catch (error) {
      opened = false;
      window.__g2bPrespecDetailHelperState.opened = false;
      window.__g2bPrespecDetailHelperState.error = String(error);
      console.error("[G2B Prespec Detail Helper] Failed to open detail.", error);
    }
  }

  const timer = window.setInterval(function () {
    openDetail();
    if (opened) {
      window.clearInterval(timer);
    }
  }, 500);

  window.setTimeout(function () {
    window.clearInterval(timer);
  }, 30000);

  openDetail();
})();
