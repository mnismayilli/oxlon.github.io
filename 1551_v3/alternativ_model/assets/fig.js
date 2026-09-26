// fig.js — Plotly loader for the 15.5.1 site (blueprint §5/§4). Usage:
//   <div class="fig" data-fig="fr01_artim"></div>
//   <script src="assets/plotly.min.js"></script>
//   <script src="assets/fig.js"></script>  (auto-loads every div.fig[data-fig] on the page)
(function () {
  var CONFIG = { displaylogo: false, responsive: true,
    modeBarButtonsToRemove: ["lasso2d", "select2d", "autoScale2d", "toggleSpikelines"] };

  function loadFig(container, figId, baseUrl) {
    var base = baseUrl || (document.currentScript && document.currentScript.src
      ? document.currentScript.src.replace(/fig\.js(\?.*)?$/, "") : "assets/");
    fetch(base + "figdata/" + figId + ".json")
      .then(function (r) { return r.json(); })
      .then(function (spec) { Plotly.newPlot(container, spec.data, spec.layout, CONFIG); })
      .catch(function (err) { container.textContent = "Figure unavailable: " + figId; console.error(err); });
  }

  function autoload() {
    document.querySelectorAll(".fig[data-fig]").forEach(function (el) {
      loadFig(el, el.getAttribute("data-fig"));
    });
  }

  window.loadFig = loadFig;
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", autoload);
  else autoload();
})();
