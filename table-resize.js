/**
 * 表格列宽手动调整，支持记住宽度
 * 使用：initResizableTable(tableId, storageKey, colWidths, opColIndex)
 */
(function () {
  window.initResizableTable = function (tableId, storageKey, colWidths, opColIndex) {
    var table = document.getElementById(tableId);
    if (!table) return;

    opColIndex = opColIndex == null ? -1 : opColIndex;
    var cols = table.querySelectorAll('colgroup col');
    if (!cols.length) return;

    function loadWidths() {
      try {
        var s = localStorage.getItem(storageKey);
        if (s) {
          var arr = JSON.parse(s);
          table.style.tableLayout = 'fixed';
          arr.forEach(function (w, i) {
            if (cols[i] && w > 0) {
              cols[i].style.width = w + 'px';
            }
          });
        } else {
          table.style.tableLayout = 'auto';
          cols.forEach(function (col) { col.style.width = ''; });
        }
      } catch (e) {}
    }

    function saveWidths() {
      try {
        var arr = [];
        ths.forEach(function (th, i) {
          var w = parseInt(cols[i].style.width, 10) || 0;
          if (w <= 0 && th.offsetWidth) w = th.offsetWidth;
          arr.push(Math.max(40, w || 80));
        });
        localStorage.setItem(storageKey, JSON.stringify(arr));
      } catch (e) {}
    }

    function setOpColHidden(hidden) {
      if (opColIndex >= 0 && cols[opColIndex]) {
        if (hidden) {
          cols[opColIndex].style.width = '0px';
          cols[opColIndex].style.minWidth = '0px';
        } else {
          try {
            var s = localStorage.getItem(storageKey);
            var w = 120;
            if (s) { var arr = JSON.parse(s); w = arr[opColIndex] > 0 ? arr[opColIndex] : 120; }
            cols[opColIndex].style.width = w + 'px';
            cols[opColIndex].style.minWidth = '';
          } catch (e) { cols[opColIndex].style.width = '120px'; }
        }
      }
    }

    var ths = table.querySelectorAll('thead th');
    ths.forEach(function (th, i) {
      if (i >= cols.length) return;
      th.style.position = 'relative';
      var handle = document.createElement('div');
      handle.className = 'col-resize-handle';
      handle.title = '拖动此处调整列宽';
      th.appendChild(handle);

      var startX, startW;
      handle.addEventListener('mousedown', function (e) {
        e.preventDefault();
        startX = e.clientX;
        var currentW = th.offsetWidth;
        startW = parseInt(cols[i].style.width, 10) || currentW || 80;
        if (startW < 20) startW = currentW || 80;

        function onMove(e) {
          var dx = e.clientX - startX;
          var newW = Math.max(40, startW + dx);
          cols[i].style.width = newW + 'px';
          table.style.tableLayout = 'fixed';
        }

        function onUp() {
          document.removeEventListener('mousemove', onMove);
          document.removeEventListener('mouseup', onUp);
          document.body.style.cursor = '';
          document.body.style.userSelect = '';
          saveWidths();
          loadWidths();
        }

        document.body.style.cursor = 'col-resize';
        document.body.style.userSelect = 'none';
        document.addEventListener('mousemove', onMove);
        document.addEventListener('mouseup', onUp);
      });
    });

    loadWidths();

    return { loadWidths: loadWidths, saveWidths: saveWidths, setOpColHidden: setOpColHidden };
  };
})();
