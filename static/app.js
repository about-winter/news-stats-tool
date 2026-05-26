/**
 * 新闻采编统计工具 - Web 版 前端逻辑
 */
(function() {
'use strict';

var STATE = { config: null, results: null, currentPage: 'summary' };
var MEDALS = ['🥇', '🥈', '🥉']; // gold, silver, bronze

// ============ 初始化 ============
document.addEventListener('DOMContentLoaded', function() {
  initNav();
  initUpload();
  loadConfig();
});

// ============ 导航 ============
function initNav() {
  document.getElementById('navMenu').addEventListener('click', function(e) {
    var btn = e.target.closest('.nav-item');
    if (!btn) return;
    switchPage(btn.dataset.page);
  });
}

function switchPage(page) {
  STATE.currentPage = page;
  document.querySelectorAll('.nav-item').forEach(function(b) {
    b.classList.toggle('active', b.dataset.page === page);
  });
  document.querySelectorAll('.page').forEach(function(p) {
    p.classList.toggle('active', p.id === 'page-' + page);
  });
}

// ============ 文件上传 ============
function initUpload() {
  document.getElementById('fileInput').addEventListener('change', function() {
    var file = this.files[0];
    if (file) uploadFile(file);
  });
}

function uploadFile(file) {
  var label = document.getElementById('uploadLabel');
  var dot = document.getElementById('fileDot');
  var nameEl = document.getElementById('fileName');

  label.textContent = '统计中...';
  label.classList.add('loading');
  dot.className = 'file-dot ready';
  nameEl.textContent = file.name;

  var fd = new FormData();
  fd.append('file', file);

  fetch('/api/upload', { method: 'POST', body: fd })
    .then(function(r) { return r.json(); })
    .then(function(data) {
      if (data.status === 'ok') {
        STATE.results = data;
        dot.className = 'file-dot done';
        label.textContent = '导入 Excel 文件';
        label.classList.remove('loading');
        document.getElementById('downloadBtn').style.display = 'flex';
        renderAll();
        switchPage('summary');
      } else {
        throw new Error(data.message || '统计失败');
      }
    })
    .catch(function(err) {
      dot.className = 'file-dot error';
      label.textContent = '导入 Excel 文件';
      label.classList.remove('loading');
      nameEl.textContent = '统计失败';
      toast(err.message || '发生未知错误', true);
    });
  this.value = '';
}

// ============ 渲染 ============
function renderAll() {
  renderBadges();
  renderAwards();
  renderSummaryTable();
  renderProgressTable();
  renderEditorTable();
}

function renderBadges() {
  var area = document.getElementById('summaryBadges');
  if (!STATE.results) {
    area.innerHTML = '<span class="badge badge-gray">等待导入文件</span>';
    return;
  }
  area.innerHTML =
    '<span class="badge badge-gray">' + STATE.results.summary_count + ' 条</span>' +
    '<span class="badge badge-success">统计完成</span>';
}

function renderAwards() {
  if (!STATE.results) return;
  var aw = STATE.results.awards;

  // 五谷丰登
  var rateHtml = '';
  aw.top_rates.forEach(function(e) {
    rateHtml += '<div class="award-item"><span class="medal">' + (MEDALS[e.rank-1]||'') + '</span> ' +
      '<span class="aname">' + esc(e.name) + '</span> ' +
      '<span class="aval">' + (e.val*100).toFixed(1) + '%</span></div>';
  });
  document.getElementById('awardRateList').innerHTML = rateHtml || '<div class="award-item"><span class="aval">暂无数据</span></div>';

  // 弹无虚发
  var hitHtml = '';
  aw.top_hits.forEach(function(e) {
    hitHtml += '<div class="award-item"><span class="medal">' + (MEDALS[e.rank-1]||'') + '</span> ' +
      '<span class="aname">' + esc(e.name) + '</span> ' +
      '<span class="aval">' + (e['已刊发数']||0) + '/' + (e['总提报']||0) + ' · ' + (e.val*100).toFixed(1) + '%</span></div>';
  });
  document.getElementById('awardHitList').innerHTML = hitHtml || '<div class="award-item"><span class="aval">暂无数据</span></div>';

  // 笔耕不辍
  var editorHtml = '';
  aw.top_editors.forEach(function(e) {
    editorHtml += '<span class="editor-badge">' + esc(e.name) + '</span> ' +
      '<span class="editor-count">' + e.count + '稿</span>';
  });
  document.getElementById('awardEditorNames').innerHTML = editorHtml;
}

// 统计汇总表格
function renderSummaryTable() {
  var body = document.getElementById('summaryTableBody');
  var pill = document.getElementById('summaryPill');
  if (!STATE.results) {
    body.innerHTML = '<div class="empty-state">请点击左侧「导入 Excel 文件」开始</div>';
    pill.style.display = 'none';
    return;
  }

  var rows = STATE.results.summary;
  var special = STATE.config ? STATE.config.special_persons : [];
  var hi = STATE.config ? STATE.config.rate_high / 100 : 0.8;
  var mi = STATE.config ? STATE.config.rate_mid / 100 : 0.5;

  pill.style.display = 'inline';
  pill.textContent = rows.length + ' 条记录';

  var h = '<table class="data-table"><thead><tr>' +
    '<th>人员</th><th>栏目</th><th>成稿/线索</th><th>新闻相关部门</th><th class="ctr">汇总数</th><th class="ctr">指标</th><th>达成率</th>' +
    '</tr></thead><tbody>';

  rows.forEach(function(r) {
    var person = esc(String(r['人员'] || ''));
    var col = esc(String(r['栏目'] || '—'));
    var typ = String(r['成稿/线索'] || '');
    var tcls = typ === '成稿' ? 'cg' : 'xs';
    var isSp = special.indexOf(String(r['人员'] || '')) >= 0;
    var dept = isSp ? esc(String(r['新闻相关部门'] || '—')) : '—';
    var dc = dept !== '—' ? 'var(--text-2)' : 'var(--text-3)';
    var cnt = r['汇总数'] != null ? Number(r['汇总数']) : 0;
    var idx = r['指标'] != null ? Number(r['指标']) : null;
    var rate = r['达成率值'] != null ? Number(r['达成率值']) : null;

    var rateHtml = '<span style="color:var(--text-3)">—</span>';
    if (rate != null && idx > 0) {
      var pct = Math.min(rate, 1);
      var cls = rate >= hi ? 'high' : (rate >= mi ? 'mid' : 'low');
      rateHtml = '<div class="rate-bar-wrap">' +
        '<div class="rate-bar"><div class="rate-bar-fill ' + cls + '" style="width:' + (pct*100) + '%"></div></div>' +
        '<span class="rate-pct ' + cls + '">' + (rate*100).toFixed(0) + '%</span></div>';
    }

    h += '<tr>' +
      '<td><div><strong>' + person + '</strong></div></td>' +
      '<td><div>' + col + '</div></td>' +
      '<td><div><span class="type-pill ' + tcls + '">' + esc(typ || '—') + '</span></div></td>' +
      '<td><div style="color:' + dc + ';font-size:11px">' + dept + '</div></td>' +
      '<td class="ctr"><div style="color:var(--blue);font-weight:bold">' + cnt + '</div></td>' +
      '<td class="ctr"><div style="color:' + (idx != null ? 'var(--text-2)' : 'var(--text-3)') + '">' + (idx != null ? idx : '—') + '</div></td>' +
      '<td><div>' + rateHtml + '</div></td>' +
      '</tr>';
  });

  h += '</tbody></table>';
  body.innerHTML = h;
}

// 达成进度表格
function renderProgressTable() {
  var body = document.getElementById('progressTableBody');
  var pill = document.getElementById('progressPill');
  if (!STATE.results) {
    body.innerHTML = '<div class="empty-state">请先导入 Excel 文件</div>';
    pill.style.display = 'none';
    return;
  }

  var rows = STATE.results.progress;
  var hi = STATE.config ? STATE.config.rate_high / 100 : 0.8;
  var mi = STATE.config ? STATE.config.rate_mid / 100 : 0.5;

  pill.style.display = 'inline';
  pill.textContent = rows.length + ' 人';

  var h = '<table class="data-table"><thead><tr>' +
    '<th>媒体通联</th><th class="ctr">指标总数</th><th class="ctr">提报总数</th><th>提报达成率</th><th class="ctr">刊发总数</th><th>刊发率</th>' +
    '</tr></thead><tbody>';

  rows.forEach(function(r) {
    var name = esc(String(r['媒体通联'] || ''));
    var idxT = Number(r['指标总数']);
    var sub = Number(r['提报总数']);
    var pub = Number(r['刊发总数']);
    var rate = Number(r['提报达成率']);
    var pubRate = Number(r['刊发率值']);

    var clsR = rate >= hi ? 'high' : (rate >= mi ? 'mid' : 'low');
    var clsP = pubRate >= hi ? 'high' : (pubRate >= mi ? 'mid' : 'low');

    h += '<tr>' +
      '<td><div><strong>' + name + '</strong></div></td>' +
      '<td class="ctr"><div style="color:var(--text-2)">' + idxT + '</div></td>' +
      '<td class="ctr"><div style="color:var(--blue);font-weight:bold">' + sub + '</div></td>' +
      '<td><div class="rate-bar-wrap">' +
        '<div class="rate-bar rate-bar-wide"><div class="rate-bar-fill ' + clsR + '" style="width:' + (Math.min(rate,1)*100) + '%"></div></div>' +
        '<span class="rate-pct ' + clsR + '">' + (rate*100).toFixed(1) + '%</span></div></td>' +
      '<td class="ctr"><div style="color:var(--text-2)">' + pub + '</div></td>' +
      '<td><div class="rate-bar-wrap">' +
        '<div class="rate-bar rate-bar-wide"><div class="rate-bar-fill ' + clsP + '" style="width:' + (Math.min(pubRate,1)*100) + '%"></div></div>' +
        '<span class="rate-pct ' + clsP + '">' + (pubRate*100).toFixed(1) + '%</span></div></td>' +
      '</tr>';
  });

  h += '</tbody></table>';
  body.innerHTML = h;
}

// 编辑统计表格
function renderEditorTable() {
  var body = document.getElementById('editorTableBody');
  var pill = document.getElementById('editorPill');
  if (!STATE.results || !STATE.results.editor || STATE.results.editor.length === 0) {
    body.innerHTML = '<div class="empty-state">请先导入 Excel 文件</div>';
    pill.style.display = 'none';
    return;
  }

  var rows = STATE.results.editor;
  var total = STATE.results.editor_total;
  var rankColors = ['gold', 'silver', 'bronze'];

  pill.style.display = 'inline';
  pill.textContent = rows.length + ' 人 / ' + total + ' 篇';

  var h = '<table class="data-table"><thead><tr><th>编辑</th><th class="ctr">发主编总数</th></tr></thead><tbody>';

  rows.forEach(function(r, i) {
    var name = esc(String(r['编辑'] || ''));
    var count = Number(r['发主编总数']);
    var rc = i < 3 ? rankColors[i] : '';
    var medal = i < 3 ? MEDALS[i] : '';
    var rankTxt = medal || (i + 1);

    h += '<tr>' +
      '<td><div><span class="editor-rank ' + rc + '">' + rankTxt + '</span> &nbsp;&nbsp;<strong>' + name + '</strong></div></td>' +
      '<td class="ctr"><div style="color:var(--blue);font-size:14px;font-weight:bold">' + count + ' 篇</div></td>' +
      '</tr>';
  });

  h += '</tbody></table>';
  body.innerHTML = h;
}

// ============ 下载 ============
window.downloadExcel = function() {
  if (!STATE.results || !STATE.results.session_id) return;
  window.open('/api/download/' + STATE.results.session_id, '_blank');
};

// ============ 设置 ============
function loadConfig() {
  fetch('/api/config')
    .then(function(r) { return r.json(); })
    .then(function(cfg) {
      STATE.config = cfg;
      document.getElementById('cfgDataSheet').value = cfg.data_sheet || '';
      document.getElementById('cfgIndexSheet').value = cfg.index_sheet || '';
      document.getElementById('cfgOutputSheet').value = cfg.output_sheet || '';
      document.getElementById('cfgSpecialPersons').value = (cfg.special_persons || []).join('、');
      document.getElementById('cfgRateHigh').value = cfg.rate_high || 80;
      document.getElementById('cfgRateMid').value = cfg.rate_mid || 50;
      document.getElementById('cfgAutoOpen').checked = !!cfg.auto_open;
    })
    .catch(function(err) { console.error('加载配置失败', err); });
}

window.saveSettings = function() {
  var body = {
    data_sheet: document.getElementById('cfgDataSheet').value.trim(),
    index_sheet: document.getElementById('cfgIndexSheet').value.trim(),
    output_sheet: document.getElementById('cfgOutputSheet').value.trim(),
    special_persons: document.getElementById('cfgSpecialPersons').value.trim(),
    rate_high: document.getElementById('cfgRateHigh').value,
    rate_mid: document.getElementById('cfgRateMid').value,
    auto_open: document.getElementById('cfgAutoOpen').checked
  };

  fetch('/api/config', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  })
    .then(function(r) { return r.json(); })
    .then(function(data) {
      if (data.status === 'ok') {
        STATE.config = data.config;
        toast('设置已保存');
        if (STATE.results) renderAll();
      } else {
        toast(data.message || '保存失败', true);
      }
    })
    .catch(function(err) { toast('保存失败: ' + err.message, true); });
};

// ============ Toast ============
var _toastTimer;
function toast(msg, isErr) {
  var el = document.getElementById('toast');
  clearTimeout(_toastTimer);
  el.textContent = msg;
  el.className = 'toast' + (isErr ? ' error' : '');
  el.classList.add('show');
  _toastTimer = setTimeout(function() { el.classList.remove('show'); }, 2500);
}

// ============ Utils ============
function esc(s) {
  if (s == null) return '';
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

})();
