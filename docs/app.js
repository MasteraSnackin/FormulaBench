const benchmarkViews = Object.freeze({
  overall: {
    label: "All 400 public tasks",
    tasks: 400,
    passed: 133,
    passRate: 33.25,
    cellRate: 40.56,
    correctCells: 120824,
    cells: 297882,
  },
  cell: {
    label: "275 cell-level tasks",
    tasks: 275,
    passed: 90,
    passRate: 32.73,
    cellRate: 67.79,
    correctCells: 10182,
    cells: 15021,
  },
  sheet: {
    label: "125 sheet-level tasks",
    tasks: 125,
    passed: 43,
    passRate: 34.4,
    cellRate: 39.11,
    correctCells: 110642,
    cells: 282861,
  },
});

const numberFormat = new Intl.NumberFormat("en-GB");

function setView(viewName) {
  const view = benchmarkViews[viewName];
  if (!view) return;

  document.querySelector("#view-label").textContent = view.label;
  document.querySelector("#passed-count").textContent = numberFormat.format(view.passed);
  document.querySelector("#view-detail").textContent = `of ${numberFormat.format(view.tasks)} tasks`;
  document.querySelector("#pass-rate").textContent = `${view.passRate.toFixed(2)}%`;
  document.querySelector("#cell-rate").textContent = `${view.cellRate.toFixed(2)}%`;
  document.querySelector("#cell-detail").textContent = `${numberFormat.format(view.correctCells)} of ${numberFormat.format(view.cells)} cells correct`;

  const passBar = document.querySelector("#pass-bar");
  const cellBar = document.querySelector("#cell-bar");
  passBar.style.setProperty("--value", `${view.passRate}%`);
  cellBar.style.setProperty("--value", `${view.cellRate}%`);
  passBar.parentElement.setAttribute("aria-label", `Task pass rate ${view.passRate.toFixed(2)} percent`);
  cellBar.parentElement.setAttribute("aria-label", `Cell accuracy ${view.cellRate.toFixed(2)} percent`);

  document.querySelectorAll("[data-view]").forEach((button) => {
    const active = button.dataset.view === viewName;
    button.classList.toggle("is-active", active);
    button.setAttribute("aria-pressed", String(active));
  });
}

document.querySelectorAll("[data-view]").forEach((button) => {
  button.addEventListener("click", () => setView(button.dataset.view));
});

const demoVideo = document.querySelector("#demo-video");
if (demoVideo) {
  demoVideo.addEventListener("error", () => {
    const message = document.querySelector("#video-status");
    if (message) message.hidden = false;
  }, true);
}
