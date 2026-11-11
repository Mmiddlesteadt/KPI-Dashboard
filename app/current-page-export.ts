/** Export the same bundled application with the current report and view state. */
export function buildCurrentPageExport(snapshot: unknown) {
  const page = document.documentElement.cloneNode(true) as HTMLElement;
  page.querySelector("#root")!.innerHTML = "";
  page.querySelectorAll("[data-wtdc-snapshot]").forEach(node => node.remove());
  const state = document.createElement("script");
  state.dataset.wtdcSnapshot = "true";
  state.textContent = "globalThis.__WTDC_SNAPSHOT__=" + JSON.stringify(snapshot).replace(/</g, "\\u003c") + ";";
  page.querySelector("head")!.appendChild(state);
  const style = document.createElement("style");
  style.textContent = ".report-actions-grid,.workspace-head>.button{display:none!important}";
  page.querySelector("head")!.appendChild(style);
  return "<!doctype html>\n" + page.outerHTML;
}
