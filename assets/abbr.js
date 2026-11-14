(function () {
  var data = window.ABBR || [];
  var inp = document.getElementById("abbrsearch");
  var host = document.getElementById("abbrlist");
  if (!inp || !host) return;
  function esc(x) {
    return String(x == null ? "" : x).replace(/[&<>"']/g, function (z) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[z];
    });
  }
  function render() {
    var q = (inp.value || "").trim().toLowerCase();
    var arr = !q ? data : data.filter(function (x) {
      return (x.a + " " + x.full + " " + x.cn + " " + x.note).toLowerCase().indexOf(q) >= 0;
    });
    host.innerHTML = arr.length ? arr.map(function (x) {
      return '<div style="padding:9px 0;border-bottom:1px dashed var(--line)">' +
        '<b style="color:#1f5c8b;font-size:15px">' + esc(x.a) + '</b> ' +
        '<span style="color:#555">' + esc(x.full) + '</span>' +
        '<div style="font-size:13px;margin-top:2px">' + esc(x.cn) + ' ｜ ' + esc(x.note) + '</div></div>';
    }).join("") : '<p class="empty">没有匹配的缩写。</p>';
  }
  inp.oninput = render;
  render();
})();
