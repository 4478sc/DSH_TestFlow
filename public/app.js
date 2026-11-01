/* ==========================================================================
   TestFlow frontend
   Vanilla JS. All data comes from the local HTTP API - there is no mock mode
   and no offline success path. The UI always reports what the server returned.
   ========================================================================== */

(function () {
  'use strict';

  var DATA_EXTS = ['.xlsx', '.xls', '.xlsm'];
  var IMAGE_EXTS = ['.jpg', '.jpeg'];
  var MAX_FILE_BYTES = 50 * 1024 * 1024;
  var INDEX_TTL = 60000;
  var RECORD_FILTER_DEFAULT = { status: 'all', mode: 'all' };

  var ESC_MAP = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };

  var state = {
    phase: 'loading',
    bootstrapping: false,
    loadError: '',
    projects: [],
    projectId: null,
    project: null,
    clipboard: { items: [], count: 0 },
    tests: [],
    summary: [],
    instructions: [],
    source: '',
    filters: { q: '', module: 'all', priority: 'all' },
    selectedId: null,
    index: new Map(),
    indexState: 'idle',
    indexFailed: new Set(),
    recordsLoadedAt: new Map(),
    recordsLoading: new Set(),
    recordErrors: new Map(),
    recordFilters: Object.assign({}, RECORD_FILTER_DEFAULT),
    form: null,
    pendingUpload: null,
    submitting: false,
    exporting: null
  };

  var inflightRecords = new Map();
  var activeModalClose = null;
  var lightboxReturnFocus = null;

  var $ = function (selector) { return document.querySelector(selector); };
  var $$ = function (selector) { return Array.prototype.slice.call(document.querySelectorAll(selector)); };

  /* ------------------------------------------------------------------------
     small utilities
     ------------------------------------------------------------------------ */

  function esc(value) {
    if (value === null || value === undefined) return '';
    return String(value).replace(/[&<>"']/g, function (ch) { return ESC_MAP[ch]; });
  }

  function text(value, fallback) {
    var s = value === null || value === undefined ? '' : String(value);
    return s === '' ? (fallback === undefined ? '' : fallback) : s;
  }

  function safeUrl(raw) {
    if (!raw) return '';
    try {
      var url = new URL(String(raw), window.location.href);
      if (url.protocol === 'http:' || url.protocol === 'https:' || url.protocol === 'blob:' || url.protocol === 'data:') {
        return url.href;
      }
    } catch (err) {
      return '';
    }
    return '';
  }

  function attachmentUrl(att) {
    if (!att) return '';
    if (att.url) return att.url;
    if (att.id) return '/api/attachments/' + att.id + '/file';
    return '';
  }

  function pad2(value) { return value < 10 ? '0' + value : String(value); }

  function todayISO() {
    var d = new Date();
    return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
  }

  function isValidISODate(value) {
    if (typeof value !== 'string' || value.length !== 10 || value.charAt(4) !== '-' || value.charAt(7) !== '-') return false;
    var y = Number(value.slice(0, 4));
    var m = Number(value.slice(5, 7));
    var d = Number(value.slice(8, 10));
    if (!isFinite(y) || !isFinite(m) || !isFinite(d)) return false;
    var dt = new Date(Date.UTC(y, m - 1, d));
    return dt.getUTCFullYear() === y && dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d;
  }

  function formatDateTime(value) {
    if (!value) return '—';
    var d = new Date(value);
    if (isNaN(d.getTime())) return String(value);
    return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()) + ' ' + pad2(d.getHours()) + ':' + pad2(d.getMinutes());
  }

  function formatBytes(value) {
    var n = Number(value);
    if (!isFinite(n) || n < 0) return '';
    if (n < 1024) return n + ' B';
    if (n < 1024 * 1024) return (n / 1024).toFixed(n < 10240 ? 1 : 0) + ' KB';
    return (n / (1024 * 1024)).toFixed(n < 10 * 1024 * 1024 ? 1 : 0) + ' MB';
  }

  function priorityRank(p) {
    var s = String(p || '').toUpperCase();
    var order = ['P0', 'P1', 'P2', 'P3'];
    var i = order.indexOf(s);
    return i < 0 ? 9 : i;
  }

  function priorityClass(p) {
    var s = String(p || '').toLowerCase();
    return s === 'p0' || s === 'p1' || s === 'p2' || s === 'p3' ? s : 'p-other';
  }

  function modeLabel(mode) { return mode === 'measurement' ? '实测' : (mode === 'simulation' ? '仿真' : '未设置'); }
  function statusLabel(status) { return status === 'Pass' ? 'Pass' : (status === 'Fail' ? 'Fail' : '未判定'); }

  function normalizeText(value) {
    return String(value === null || value === undefined ? '' : value).toLowerCase();
  }

  function normalizeRecord(raw) {
    var record = raw && typeof raw === 'object' ? Object.assign({}, raw) : {};
    record.id = Number(record.id);
    record.attachments = Array.isArray(record.attachments) ? record.attachments.map(normalizeAttachment) : [];
    record.mode = record.mode === 'measurement' ? 'measurement' : (record.mode === 'simulation' ? 'simulation' : '');
    record.status = record.status === 'Pass' || record.status === 'Fail' ? record.status : '';
    record.test_date = record.test_date || '';
    record.result = record.result || '';
    record.unit = record.unit || '';
    record.notes = record.notes || '';
    return record;
  }

  function normalizeAttachment(raw) {
    var att = raw && typeof raw === 'object' ? Object.assign({}, raw) : {};
    att.id = Number(att.id);
    att.size = Number(att.size) || 0;
    att.name = att.name || '附件';
    att.kind = att.kind === 'image' ? 'image' : 'data';
    return att;
  }

  function extensionOf(name) {
    var s = String(name || '');
    var dot = s.lastIndexOf('.');
    return dot < 0 ? '' : s.slice(dot).toLowerCase();
  }

  function fileNameIsData(name) { return DATA_EXTS.indexOf(extensionOf(name)) >= 0; }
  function fileNameIsImage(name) { return IMAGE_EXTS.indexOf(extensionOf(name)) >= 0; }

  function isTypingTarget(target) {
    if (!target || !target.tagName) return false;
    var tag = target.tagName.toLowerCase();
    return tag === 'input' || tag === 'textarea' || tag === 'select' || target.isContentEditable;
  }

  function selectedTest() {
    for (var i = 0; i < state.tests.length; i++) {
      if (state.tests[i].id === state.selectedId) return state.tests[i];
    }
    return null;
  }

  function findRecord(testId, recordId) {
    var list = state.index.get(testId) || [];
    for (var i = 0; i < list.length; i++) {
      if (list[i].id === recordId) return list[i];
    }
    return null;
  }

  function aggregateRecords() {
    var agg = { total: 0, pass: 0, fail: 0, blank: 0, simulation: 0, measurement: 0, projects: 0, attachments: 0 };
    state.index.forEach(function (records) {
      if (records.length) agg.projects += 1;
      for (var i = 0; i < records.length; i++) {
        var r = records[i];
        agg.total += 1;
        if (r.status === 'Pass') agg.pass += 1;
        else if (r.status === 'Fail') agg.fail += 1;
        else agg.blank += 1;
        if (r.mode === 'simulation') agg.simulation += 1;
        if (r.mode === 'measurement') agg.measurement += 1;
        agg.attachments += (r.attachments || []).length;
      }
    });
    return agg;
  }

  function moduleNameOf(test) { return text(test && test.module, '未分模块'); }

  function sortedPriorities() {
    var values = [];
    state.tests.forEach(function (test) {
      var p = text(test.priority);
      if (p && values.indexOf(p) < 0) values.push(p);
    });
    values.sort(function (a, b) { return priorityRank(a) - priorityRank(b) || String(a).localeCompare(String(b), 'zh-CN'); });
    return values;
  }

  /* ------------------------------------------------------------------------
     HTTP layer
     ------------------------------------------------------------------------ */

  class RequestError extends Error {
    constructor(message, status, kind) {
      super(message);
      this.name = 'RequestError';
      this.status = status || 0;
      this.kind = kind || 'http';
    }
  }

  async function apiJson(path, options) {
    options = options || {};
    var headers = Object.assign({}, options.headers || {});
    var body = options.body;
    if (Object.prototype.hasOwnProperty.call(options, 'json')) {
      headers['Content-Type'] = 'application/json';
      body = JSON.stringify(options.json);
    }
    var controller = new AbortController();
    var timeout = options.timeout || 30000;
    var timer = setTimeout(function () { controller.abort(); }, timeout);
    try {
      var res = await fetch(path, {
        method: options.method || 'GET',
        headers: headers,
        body: body,
        credentials: 'same-origin',
        cache: 'no-store',
        signal: controller.signal
      });
      var data = null;
      var ctype = res.headers.get('content-type') || '';
      if (ctype.indexOf('application/json') >= 0) {
        try { data = await res.json(); } catch (err) { data = null; }
      }
      if (!res.ok) {
        var message = data && data.error ? String(data.error) : '请求失败（HTTP ' + res.status + '）';
        throw new RequestError(message, res.status, 'http');
      }
      return data;
    } catch (err) {
      if (err && err.name === 'AbortError') {
        throw new RequestError('请求超时，请检查本地服务是否响应。', 0, 'timeout');
      }
      if (err instanceof RequestError) throw err;
      throw new RequestError('无法连接本地服务，请确认 server.py 正在运行。', 0, 'network');
    } finally {
      clearTimeout(timer);
    }
  }

  function projectApiPath(suffix) {
    return '/api/projects/' + encodeURIComponent(state.projectId || '') + suffix;
  }

  function projectBootstrapPath() {
    return state.projectId ? projectApiPath('/bootstrap') : '/api/bootstrap';
  }

  function currentProject() {
    if (state.project) return state.project;
    return state.projects.filter(function (p) { return p.id === state.projectId; })[0] || null;
  }

  function projectFromUrl() {
    try {
      return new URLSearchParams(window.location.search).get('project') || '';
    } catch (err) {
      return '';
    }
  }

  function syncProjectUrl() {
    if (!window.history || !window.history.replaceState) return;
    try {
      var url = new URL(window.location.href);
      if (state.projectId) url.searchParams.set('project', state.projectId);
      else url.searchParams.delete('project');
      window.history.replaceState(null, '', url.toString());
    } catch (err) { /* ignore */ }
  }

  function errorMessage(err) {
    if (!err) return '未知错误';
    return err.message || String(err);
  }

  function ambiguousError(err) {
    return !!err && (err.kind === 'timeout' || err.kind === 'network' || err.kind === 'bad-response' || err.status >= 500);
  }

  function fetchRecords(testId, options) {
    if (inflightRecords.has(testId)) return inflightRecords.get(testId);
    var promise = apiJson(projectApiPath('/tests/' + encodeURIComponent(testId) + '/records'), { timeout: 30000 }).then(function (data) {
      if (!data || !Array.isArray(data.records)) {
        throw new RequestError('服务器返回的记录数据格式异常。', 0, 'bad-response');
      }
      return data.records.map(normalizeRecord);
    });
    var wrapped = promise.then(function (value) {
      inflightRecords.delete(testId);
      return value;
    }, function (err) {
      inflightRecords.delete(testId);
      throw err;
    });
    inflightRecords.set(testId, wrapped);
    return wrapped;
  }

  async function uploadFiles(recordId, group) {
    var form = new FormData();
    group.files.forEach(function (file) { form.append('files', file, file.name); });
    form.append('kind', group.kind);
    var data = await apiJson('/api/records/' + encodeURIComponent(recordId) + '/attachments', {
      method: 'POST',
      body: form,
      timeout: 180000
    });
    if (!data || !Array.isArray(data.attachments)) {
      throw new RequestError('服务器未返回已上传的附件信息。', 0, 'bad-response');
    }
    return data.attachments.map(normalizeAttachment);
  }

  async function downloadBlob(url, fallbackName) {
    var controller = new AbortController();
    var timer = setTimeout(function () { controller.abort(); }, 180000);
    try {
      var res = await fetch(url, { credentials: 'same-origin', cache: 'no-store', signal: controller.signal });
      if (!res.ok) {
        var message = '下载失败（HTTP ' + res.status + '）';
        var ctype = res.headers.get('content-type') || '';
        if (ctype.indexOf('application/json') >= 0) {
          try {
            var data = await res.json();
            if (data && data.error) message = String(data.error);
          } catch (err) { /* keep status message */ }
        }
        throw new RequestError(message, res.status, 'http');
      }
      var blob = await res.blob();
      if (!blob || !blob.size) throw new RequestError('服务器返回了空文件。', 0, 'bad-response');
      var name = filenameFromDisposition(res.headers.get('content-disposition')) || fallbackName;
      var objectUrl = URL.createObjectURL(blob);
      var link = document.createElement('a');
      link.href = objectUrl;
      link.download = name;
      document.body.appendChild(link);
      link.click();
      link.remove();
      setTimeout(function () { URL.revokeObjectURL(objectUrl); }, 60000);
    } catch (err) {
      if (err && err.name === 'AbortError') {
        throw new RequestError('下载超时，请稍后重试。', 0, 'timeout');
      }
      if (err instanceof RequestError) throw err;
      throw new RequestError('无法连接本地服务，下载未完成。', 0, 'network');
    } finally {
      clearTimeout(timer);
    }
  }

  function filenameFromDisposition(value) {
    if (!value) return '';
    var lower = value.toLowerCase();
    var marker = "filename*=utf-8''";
    var at = lower.indexOf(marker);
    if (at >= 0) {
      var star = value.slice(at + marker.length);
      var semi = star.indexOf(';');
      if (semi >= 0) star = star.slice(0, semi);
      star = star.trim();
      if (star.length >= 2 && star.charAt(0) === '"' && star.charAt(star.length - 1) === '"') star = star.slice(1, -1);
      try { return decodeURIComponent(star); } catch (err) { return star; }
    }
    at = lower.indexOf('filename=');
    if (at < 0) return '';
    var name = value.slice(at + 9).trim();
    semi = name.indexOf(';');
    if (semi >= 0) name = name.slice(0, semi);
    name = name.trim();
    if (name.length >= 2 && name.charAt(0) === '"' && name.charAt(name.length - 1) === '"') name = name.slice(1, -1);
    return name;
  }

  /* ------------------------------------------------------------------------
     bootstrap and record index
     ------------------------------------------------------------------------ */

  async function loadBootstrap() {
    if (state.bootstrapping) return;
    if (!state.projectId) {
      state.phase = 'ready';
      state.tests = [];
      state.summary = [];
      state.instructions = [];
      state.source = '';
      renderAll();
      renderDetail();
      return;
    }
    state.bootstrapping = true;
    state.phase = 'loading';
    state.loadError = '';
    state.selectedId = null;
    state.form = null;
    state.pendingUpload = null;
    state.index = new Map();
    state.indexState = 'idle';
    state.indexFailed = new Set();
    state.recordsLoadedAt = new Map();
    state.recordsLoading = new Set();
    state.recordErrors = new Map();
    state.recordFilters = Object.assign({}, RECORD_FILTER_DEFAULT);
    inflightRecords = new Map();
    state.filters = { q: '', module: 'all', priority: 'all' };
    var search = $('#searchInput');
    if (search) search.value = '';

    renderInitial();

    try {
      var data = await apiJson(projectBootstrapPath(), { timeout: 25000 });
      if (!data || !Array.isArray(data.tests)) {
        throw new RequestError('服务器返回的测试项目数据格式异常。', 0, 'bad-response');
      }
      state.tests = data.tests.filter(function (test) {
        return test && typeof test === 'object' && typeof test.id === 'string' && test.id !== '';
      });
      state.summary = Array.isArray(data.summary) ? data.summary.filter(function (row) { return row && typeof row === 'object'; }) : [];
      state.instructions = Array.isArray(data.instructions) ? data.instructions : [];
      state.source = typeof data.source === 'string' ? data.source : '';
      if (data.project && typeof data.project === 'object') state.project = data.project;
      state.phase = 'ready';
      state.bootstrapping = false;

      renderAll();
      loadRecordIndex();

      var first = filteredTests()[0];
      var wide = window.matchMedia('(min-width: 901px)').matches;
      if (first && wide) {
        state.selectedId = first.id;
        document.body.classList.add('mobile-detail');
        renderDetail(true);
        ensureRecordsForSelected(first.id);
      } else {
        renderDetail();
      }
    } catch (err) {
      state.phase = 'error';
      state.loadError = '无法加载测试项目：' + errorMessage(err);
      state.bootstrapping = false;
      renderBootstrapError();
    }
  }

  async function loadRecordIndex(options) {
    options = options || {};
    if (state.phase !== 'ready' || !state.tests.length) return;
    var onlyFailed = !!options.onlyFailed && state.indexFailed.size > 0;
    var targets = onlyFailed ? Array.from(state.indexFailed) : state.tests.map(function (test) { return test.id; });
    if (!targets.length) return;

    var failed = new Set();
    if (onlyFailed) {
      state.indexFailed.forEach(function (id) { if (targets.indexOf(id) < 0) failed.add(id); });
    }
    state.indexState = 'loading';
    renderTopStats();
    renderRecordStats();
    renderTestList();

    await runPool(targets, 4, async function (id) {
      try {
        var records = await fetchRecords(id);
        state.index.set(id, records);
        state.recordsLoadedAt.set(id, Date.now());
        state.indexFailed.delete(id);
        state.recordErrors.delete(id);
      } catch (err) {
        failed.add(id);
        state.indexFailed.add(id);
        if (!state.index.has(id)) state.recordErrors.set(id, errorMessage(err));
      }
    });

    failed.forEach(function (id) { if (state.index.has(id)) failed.delete(id); });
    state.indexFailed = failed;
    state.indexState = failed.size === 0 ? 'ready' : (state.index.size ? 'partial' : 'error');
    renderTopStats();
    renderRecordStats();
    renderTestList();
    if (state.selectedId) {
      renderDetailStats();
      renderRecordsArea();
    }
  }

  async function runPool(items, limit, worker) {
    var queue = items.slice();
    var workers = [];
    var count = Math.min(limit, queue.length);
    for (var i = 0; i < count; i++) {
      workers.push((async function () {
        while (queue.length) {
          var item = queue.shift();
          await worker(item);
        }
      })());
    }
    await Promise.all(workers);
  }

  function ensureRecordsForSelected(testId) {
    if (!testId) return;
    var cached = state.index.get(testId);
    var loadedAt = state.recordsLoadedAt.get(testId) || 0;
    if (cached && Date.now() - loadedAt < INDEX_TTL) return;
    if (state.recordsLoading.has(testId)) return;
    refreshRecords(testId, { force: false, notify: false }).catch(function () { /* rendered inline */ });
  }

  async function refreshRecords(testId, options) {
    options = options || {};
    if (state.recordsLoading.has(testId)) return null;
    state.recordsLoading.add(testId);
    state.recordErrors.delete(testId);
    if (state.selectedId === testId) {
      renderDetailStats();
      renderRecordsArea();
    }
    try {
      var records = await fetchRecords(testId, { force: !!options.force });
      state.index.set(testId, records);
      state.recordsLoadedAt.set(testId, Date.now());
      state.indexFailed.delete(testId);
      state.recordErrors.delete(testId);
      if (state.indexState !== 'loading') {
        state.indexState = state.indexFailed.size === 0 ? 'ready' : (state.index.size ? 'partial' : 'error');
      }
      if (state.selectedId === testId) renderIndexDerived();
      return records;
    } catch (err) {
      state.recordErrors.set(testId, errorMessage(err));
      state.indexFailed.add(testId);
      if (state.indexState !== 'loading') {
        state.indexState = state.index.size ? 'partial' : 'error';
      }
      if (options.notify) showToast('记录读取失败：' + errorMessage(err), 'error', 6500);
      throw err;
    } finally {
      state.recordsLoading.delete(testId);
      if (state.selectedId === testId) {
        renderDetailStats();
        renderRecordsArea();
      }
    }
  }

  function mergeRecord(testId, record) {
    var list = (state.index.get(testId) || []).slice();
    var index = -1;
    for (var i = 0; i < list.length; i++) {
      if (list[i].id === record.id) { index = i; break; }
    }
    if (index >= 0) {
      var previous = list[index];
      list[index] = Object.assign({}, previous, record, {
        attachments: Array.isArray(record.attachments) ? record.attachments : (previous.attachments || [])
      });
    } else {
      list.push(record);
    }
    list.sort(function (a, b) { return a.id - b.id; });
    state.index.set(testId, list);
    state.recordsLoadedAt.set(testId, Date.now());
    state.indexFailed.delete(testId);
    state.recordErrors.delete(testId);
    if (state.indexState !== 'loading') {
      state.indexState = state.indexFailed.size === 0 ? 'ready' : (state.index.size ? 'partial' : 'error');
    }
    renderIndexDerived();
  }

  function mergeAttachments(testId, recordId, attachments) {
    var record = findRecord(testId, recordId);
    if (!record) return;
    var existing = (record.attachments || []).slice();
    attachments.forEach(function (att) {
      var found = false;
      for (var i = 0; i < existing.length; i++) {
        if (existing[i].id === att.id) { existing[i] = att; found = true; break; }
      }
      if (!found) existing.push(att);
    });
    record.attachments = existing;
    renderIndexDerived();
    if (state.selectedId === testId) {
      renderDetailStats();
      renderRecordsArea();
    }
  }

  function removeAttachmentLocal(testId, recordId, attachmentId) {
    var record = findRecord(testId, recordId);
    if (!record) return;
    record.attachments = (record.attachments || []).filter(function (att) { return att.id !== attachmentId; });
    renderIndexDerived();
    if (state.selectedId === testId) renderRecordsArea();
  }

  function renderIndexDerived() {
    renderTopStats();
    renderRecordStats();
    renderTestList();
  }

  /* ------------------------------------------------------------------------
     top bar / sidebar / list rendering
     ------------------------------------------------------------------------ */

  function topStat(label, value, modifier, title) {
    return '<div class="top-stat ' + (modifier || '') + '"' + (title ? ' title="' + esc(title) + '"' : '') + '>' +
      '<span class="top-stat-label">' + esc(label) + '</span>' +
      '<span class="top-stat-value">' + esc(value) + '</span>' +
      '</div>';
  }

  function renderTopStats() {
    var host = $('#topStats');
    if (!host) return;
    if (state.phase === 'loading') {
      host.innerHTML = topStat('测试项', '…') + topStat('模块', '…') + topStat('记录', '…');
      return;
    }
    if (state.phase !== 'ready') {
      host.innerHTML = topStat('测试项', '—') + topStat('模块', '—') + topStat('记录', '—');
      return;
    }
    var p0Count = state.tests.filter(function (test) { return String(test.priority || '').toUpperCase() === 'P0'; }).length;
    var modules = [];
    state.tests.forEach(function (test) {
      var m = moduleNameOf(test);
      if (modules.indexOf(m) < 0) modules.push(m);
    });

    var recordValue = '—';
    var recordModifier = '';
    var recordTitle = '';
    if (state.indexState === 'loading') {
      recordValue = '…';
    } else if (state.index.size > 0) {
      var agg = aggregateRecords();
      recordValue = String(agg.total) + (state.indexFailed.size ? '+' : '');
      if (state.indexFailed.size) {
        recordModifier = 'is-warn';
        recordTitle = '仍有 ' + state.indexFailed.size + ' 个项目的记录未能读取，统计可能不完整。';
      }
    } else if (state.indexState === 'error') {
      recordValue = '—';
      recordModifier = 'is-warn';
      recordTitle = '记录统计读取失败。';
    } else {
      recordValue = '0';
    }

    host.innerHTML = topStat('测试项', String(state.tests.length)) +
      topStat('模块', String(modules.length)) +
      topStat('P0 项目', String(p0Count), 'is-accent') +
      topStat('记录', recordValue, recordModifier, recordTitle);
  }

  function renderModuleList() {
    var host = $('#moduleList');
    var total = $('#moduleTotal');
    if (!host) return;
    if (state.phase === 'loading') {
      host.innerHTML = '';
      for (var i = 0; i < 6; i++) {
        host.innerHTML += '<div class="skeleton" style="height:29px;margin:2px 0"></div>';
      }
      if (total) total.textContent = '';
      return;
    }
    if (state.phase !== 'ready') {
      host.innerHTML = '';
      if (total) total.textContent = '';
      return;
    }

    var modules = [];
    var counts = {};
    state.tests.forEach(function (test) {
      var m = moduleNameOf(test);
      if (!Object.prototype.hasOwnProperty.call(counts, m)) { counts[m] = 0; modules.push(m); }
      counts[m] += 1;
    });

    var parts = [];
    parts.push(moduleButton('all', '全部测试', state.tests.length, state.filters.module === 'all'));
    modules.forEach(function (name) {
      parts.push(moduleButton(name, name, counts[name], state.filters.module === name));
    });
    host.innerHTML = parts.join('');
    if (total) total.textContent = modules.length + ' 个';
  }

  function moduleButton(value, label, count, active) {
    return '<button type="button" class="module-btn' + (active ? ' active' : '') + '" data-module="' + esc(value) + '"' +
      ' aria-pressed="' + (active ? 'true' : 'false') + '">' +
      '<span class="module-dot"></span>' +
      '<span class="module-name">' + esc(label) + '</span>' +
      '<span class="module-count">' + count + '</span>' +
      '</button>';
  }

  function renderPriorityStats() {
    var host = $('#priorityStats');
    if (!host) return;
    if (state.phase === 'loading') {
      host.innerHTML = '<div class="skeleton" style="height:10px"></div><div class="skeleton" style="height:10px"></div><div class="skeleton" style="height:10px"></div>';
      return;
    }
    if (state.phase !== 'ready' || !state.tests.length) {
      host.innerHTML = '<p class="stats-note">暂无优先级数据。</p>';
      return;
    }
    var counts = {};
    state.tests.forEach(function (test) {
      var p = text(test.priority, '未标注');
      counts[p] = (counts[p] || 0) + 1;
    });
    var priorities = Object.keys(counts).sort(function (a, b) { return priorityRank(a) - priorityRank(b) || String(a).localeCompare(String(b), 'zh-CN'); });
    var max = 1;
    priorities.forEach(function (p) { if (counts[p] > max) max = counts[p]; });
    var parts = [];
    priorities.forEach(function (p) {
      var pct = Math.max(3, Math.round((counts[p] / max) * 100));
      parts.push('<div class="prio-row ' + priorityClass(p) + '">' +
        '<span class="prio-label">' + esc(p) + '</span>' +
        '<span class="prio-bar"><span class="prio-bar-fill" style="width:' + pct + '%"></span></span>' +
        '<span class="prio-num">' + counts[p] + '</span>' +
        '</div>');
    });
    host.innerHTML = parts.join('');
  }

  function renderRecordStats() {
    var host = $('#recordStats');
    var retry = $('#retryIndexBtn');
    if (!host) return;
    if (retry) {
      var canRetry = state.phase === 'ready' && (state.indexState === 'partial' || state.indexState === 'error');
      retry.hidden = !canRetry;
      retry.textContent = state.indexFailed.size ? '重试失败项' : '重试';
    }
    if (state.phase === 'loading') {
      host.innerHTML = '<div class="stats-loading">' +
        '<div class="skeleton" style="height:38px"></div>' +
        '<div class="skeleton" style="height:11px"></div>' +
        '<div class="skeleton" style="height:11px"></div>' +
        '<div class="skeleton" style="height:11px"></div>' +
        '</div>';
      return;
    }
    if (state.phase !== 'ready') {
      host.innerHTML = '<p class="stats-note">无法读取记录统计。</p>';
      return;
    }
    if (!state.tests.length) {
      host.innerHTML = '<p class="stats-note">暂无测试项目。</p>';
      return;
    }
    if (state.indexState === 'loading' && state.index.size === 0) {
      host.innerHTML = '<div class="stats-loading">' +
        '<div class="skeleton" style="height:38px"></div>' +
        '<div class="skeleton" style="height:11px"></div>' +
        '<div class="skeleton" style="height:11px"></div>' +
        '<div class="skeleton" style="height:11px"></div>' +
        '</div>';
      return;
    }
    if (state.indexState === 'error' && state.index.size === 0) {
      host.innerHTML = '<div class="stats-error">记录统计暂不可用。请确认后端服务正常，或点击上方“重试”。</div>';
      return;
    }

    var agg = aggregateRecords();
    var partial = state.indexFailed.size > 0 || state.indexState === 'loading';
    var suffix = partial ? '+' : '';
    var parts = [];
    parts.push('<div class="mini-grid">' +
      '<div class="mini-stat"><b>' + agg.total + suffix + '</b><span>记录总数</span></div>' +
      '<div class="mini-stat"><b>' + agg.projects + '</b><span>已登记项目</span></div>' +
      '</div>');
    parts.push('<div class="stat-line"><span class="dot pass"></span>Pass<b>' + agg.pass + suffix + '</b></div>');
    parts.push('<div class="stat-line"><span class="dot fail"></span>Fail<b>' + agg.fail + suffix + '</b></div>');
    parts.push('<div class="stat-line"><span class="dot blank"></span>未判定<b>' + agg.blank + suffix + '</b></div>');
    parts.push('<div class="stat-line"><span class="dot sim"></span>仿真<b>' + agg.simulation + suffix + '</b></div>');
    parts.push('<div class="stat-line"><span class="dot meas"></span>实测<b>' + agg.measurement + suffix + '</b></div>');
    if (state.indexState === 'loading') {
      parts.push('<p class="stats-note">正在读取其余项目的记录，当前数字可能继续增加。</p>');
    } else if (partial) {
      parts.push('<p class="stats-note">已读取 ' + state.index.size + ' / ' + state.tests.length + ' 个项目的记录；' +
        state.indexFailed.size + ' 个项目读取失败，统计不完整。</p>');
    }
    host.innerHTML = parts.join('');
  }

  function renderInstructions() {
    var host = $('#instructionList');
    var count = $('#instructionCount');
    if (!host) return;
    if (count) count.textContent = state.instructions.length ? String(state.instructions.length) : '';
    if (state.phase === 'loading') {
      host.innerHTML = '<div class="skeleton" style="height:52px"></div>';
      return;
    }
    if (!state.instructions.length) {
      host.innerHTML = '<p class="stats-note">暂无填写说明。</p>';
      return;
    }
    var parts = [];
    state.instructions.forEach(function (item) {
      var title = '说明';
      var body = '';
      if (typeof item === 'string') {
        body = item;
      } else if (item && typeof item === 'object') {
        var keys = Object.keys(item);
        title = text(item['项目'] || item.title || item.name || item['主题'], '说明');
        body = text(item['说明'] || item.text || item.description || item['内容']);
        if (!body) {
          var lines = [];
          keys.forEach(function (key) {
            if (key === '项目' || key === 'title' || key === 'name') return;
            var v = item[key];
            if (v !== null && v !== undefined && String(v) !== '') lines.push(key + '：' + String(v));
          });
          body = lines.join('；');
        }
      }
      if (!body) return;
      parts.push('<div class="instruction-item"><h4>' + esc(title) + '</h4><p>' + esc(body) + '</p></div>');
    });
    host.innerHTML = parts.length ? parts.join('') : '<p class="stats-note">暂无填写说明。</p>';
  }

  function renderSource() {
    var host = $('#sourceLine');
    if (!host) return;
    if (state.phase !== 'ready') {
      host.innerHTML = '<span class="source-text">数据源 · 读取中…</span>';
      return;
    }
    var source = text(state.source, '本地工作区数据');
    host.innerHTML = '<span>数据源</span><span class="source-text" title="' + esc(source) + '">' + esc(source) + '</span>';
  }

  function renderPriorityFilter() {
    var select = $('#priorityFilter');
    if (!select) return;
    var values = sortedPriorities();
    var counts = {};
    state.tests.forEach(function (test) {
      var p = text(test.priority);
      if (p) counts[p] = (counts[p] || 0) + 1;
    });
    var parts = ['<option value="all">全部</option>'];
    values.forEach(function (p) {
      parts.push('<option value="' + esc(p) + '">' + esc(p) + ' · ' + (counts[p] || 0) + '</option>');
    });
    select.innerHTML = parts.join('');
    if (state.filters.priority !== 'all' && values.indexOf(state.filters.priority) < 0) {
      state.filters.priority = 'all';
    }
    select.value = state.filters.priority;
  }

  function searchHaystack(test) {
    return normalizeText([
      test.id, test.priority, test.module, test.title, test.purpose,
      test.stimulus, test.metrics, test.conditions, test.calculation,
      test.criteria, test.instruments, test.source_notes
    ].join(' '));
  }

  function filteredTests() {
    var q = state.filters.q.trim().toLowerCase();
    var tokens = q ? q.split(' ').filter(function (token) { return token !== ''; }) : [];
    return state.tests.filter(function (test) {
      if (state.filters.module !== 'all' && moduleNameOf(test) !== state.filters.module) return false;
      if (state.filters.priority !== 'all' && text(test.priority) !== state.filters.priority) return false;
      if (tokens.length) {
        var hay = searchHaystack(test);
        for (var i = 0; i < tokens.length; i++) {
          if (hay.indexOf(tokens[i]) < 0) return false;
        }
      }
      return true;
    });
  }

  function renderTestList() {
    var host = $('#testList');
    var count = $('#testCount');
    if (!host) return;

    if (state.phase === 'loading') {
      if (count) count.textContent = '—';
      var skeletons = [];
      for (var s = 0; s < 6; s++) {
        skeletons.push('<div class="skeleton-row">' +
          '<div class="skeleton line-a"></div>' +
          '<div class="skeleton line-b"></div>' +
          '<div class="skeleton line-c"></div>' +
          '</div>');
      }
      host.innerHTML = '<div class="list-skeleton">' + skeletons.join('') + '</div>';
      renderFilterChrome();
      return;
    }

    if (state.phase !== 'ready') {
      if (count) count.textContent = '—';
      host.innerHTML = '<div class="pane-empty">' +
        '<div class="empty-icon">' + icon('alert') + '</div>' +
        '<h3>测试列表加载失败</h3>' +
        '<p>请确认本地服务运行正常，然后重试加载。</p>' +
        '<button type="button" class="btn btn-sm" data-action="retry-bootstrap">重新加载</button>' +
        '</div>';
      renderFilterChrome();
      return;
    }

    var tests = filteredTests();
    if (count) count.textContent = tests.length === state.tests.length ? String(tests.length) : (tests.length + ' / ' + state.tests.length);

    if (!state.tests.length) {
      host.innerHTML = '<div class="pane-empty">' +
        '<div class="empty-icon">' + icon('database') + '</div>' +
        '<h3>暂无测试项目</h3>' +
        '<p>请检查 seed.json 是否已被后端成功导入，然后刷新页面。</p>' +
        '</div>';
      renderFilterChrome();
      return;
    }

    if (!tests.length) {
      host.innerHTML = '<div class="pane-empty">' +
        '<div class="empty-icon">' + icon('search') + '</div>' +
        '<h3>没有匹配的测试项</h3>' +
        '<p>尝试更换关键词、模块或优先级。</p>' +
        '<button type="button" class="btn btn-sm" data-action="clear-all-filters">清除筛选</button>' +
        '</div>';
      renderFilterChrome();
      return;
    }

    var parts = [];
    tests.forEach(function (test) {
      parts.push(testCard(test));
    });
    host.innerHTML = parts.join('');
    renderFilterChrome();
  }

  function testCard(test) {
    var records = state.index.get(test.id);
    var recordInfo = '';
    if (records) {
      var latest = records.length ? records[records.length - 1] : null;
      var dotClass = latest ? (latest.status === 'Pass' ? 'pass' : (latest.status === 'Fail' ? 'fail' : '')) : '';
      recordInfo = '<span class="test-side">' + records.length + ' 条' +
        (latest ? ' <i class="status-dot ' + dotClass + '" title="最近记录：' + esc(statusLabel(latest.status)) + '"></i>' : '') +
        '</span>';
    }
    var active = state.selectedId === test.id;
    return '<button type="button" class="test-card' + (active ? ' active' : '') + '" data-test-id="' + esc(test.id) + '"' +
      ' aria-pressed="' + (active ? 'true' : 'false') + '">' +
      '<span class="test-row">' +
      '<span class="test-id">' + esc(test.id) + '</span>' +
      '<span class="badge prio-badge ' + priorityClass(test.priority) + '">' + esc(text(test.priority, '—')) + '</span>' +
      '<span class="test-module">' + esc(moduleNameOf(test)) + '</span>' +
      recordInfo +
      '</span>' +
      '<span class="test-title">' + esc(text(test.title, '未命名测试项')) + '</span>' +
      '<span class="test-purpose">' + esc(text(test.purpose, '未填写测试目的')) + '</span>' +
      '</button>';
  }

  function renderFilterChrome() {
    var clear = $('#clearFiltersBtn');
    var activeHost = $('#activeFilters');
    var hasFilter = state.filters.q.trim() !== '' || state.filters.module !== 'all' || state.filters.priority !== 'all';
    if (clear) clear.hidden = !hasFilter;
    if (!activeHost) return;
    if (!hasFilter) {
      activeHost.innerHTML = '';
      return;
    }
    var parts = [];
    if (state.filters.q.trim()) {
      parts.push(filterChip('搜索', state.filters.q.trim(), 'q'));
    }
    if (state.filters.module !== 'all') {
      parts.push(filterChip('模块', state.filters.module, 'module'));
    }
    if (state.filters.priority !== 'all') {
      parts.push(filterChip('优先级', state.filters.priority, 'priority'));
    }
    activeHost.innerHTML = parts.join('');
  }

  function filterChip(label, value, key) {
    return '<span class="filter-chip">' + esc(label) + '：' + esc(value) +
      '<button type="button" data-clear-filter="' + esc(key) + '" aria-label="移除' + esc(label) + '筛选">' + icon('close') + '</button>' +
      '</span>';
  }

  /* ------------------------------------------------------------------------
     detail rendering
     ------------------------------------------------------------------------ */

  function icon(name, extraClass) {
    return '<svg class="icon' + (extraClass ? ' ' + extraClass : '') + '" aria-hidden="true"><use href="#i-' + name + '"></use></svg>';
  }

  function renderDetail(resetScroll) {
    var pane = $('#detailPane');
    if (!pane) return;
    var top = resetScroll ? 0 : pane.scrollTop;

    if (state.phase === 'loading') {
      pane.innerHTML = '<div class="detail-inner"><div class="detail-skeleton">' +
        '<div class="skeleton skeleton-title"></div>' +
        '<div class="skeleton skeleton-sub"></div>' +
        '<div class="skeleton skeleton-block"></div>' +
        '<div class="skeleton skeleton-row-wide"></div>' +
        '<div class="skeleton skeleton-row-wide"></div>' +
        '</div></div>';
      return;
    }

    if (state.phase !== 'ready') {
      pane.innerHTML = '<div class="detail-inner"><div class="fatal-error">' +
        '<div class="error-badge">' + icon('alert') + '</div>' +
        '<h2>无法加载测试项目</h2>' +
        '<p>' + esc(state.loadError || '本地服务没有返回可用数据。') + '</p>' +
        '<p>请确认 TestFlow 后端已启动，并使用当前地址重新打开本页面。</p>' +
        '<button type="button" class="btn btn-primary" data-action="retry-bootstrap">' + icon('refresh') + '重新加载</button>' +
        '</div></div>';
      return;
    }

    var test = selectedTest();
    if (!test) {
      pane.innerHTML = '<div class="detail-inner"><div class="pane-empty" style="padding-top:80px">' +
        '<div class="empty-icon">' + icon('database') + '</div>' +
        '<h3>' + (state.projects.length ? (state.tests.length ? '选择一个测试项目' : '该项目还没有测试项') : '还没有任何项目') + '</h3>' +
        '<p>' + (state.projects.length
          ? (state.tests.length ? '从左侧列表选择项目，查看测试指导、历史记录并登记新的仿真或实测结果。' : '点击左侧“新建测试项”添加，或在顶栏“+”中上传表格新建项目。')
          : '点击顶栏的“+”，上传一份测试计划表格（.xlsx / .xlsm）创建第一个项目。') + '</p>' +
        '</div></div>';
      return;
    }

    var records = state.index.get(test.id) || [];
    pane.innerHTML = '<div class="detail-inner">' +
      '<button type="button" class="btn btn-quiet btn-sm mobile-back" data-action="back-to-list">' + icon('arrow-left') + '返回测试列表</button>' +
      detailHeadHTML(test, records) +
      '<div id="detailAlert">' + detailAlertHTML() + '</div>' +
      guidanceHTML(test) +
      sourcePanelHTML(test) +
      '<div class="form-host" id="formHost">' + (state.form ? formHTML() : '') + '</div>' +
      '<section class="records-section" id="recordsArea">' + recordsAreaHTML(test) + '</section>' +
      '</div>';

    if (!resetScroll) pane.scrollTop = top;
    wireImageFallbacks();
  }

  function detailHeadHTML(test, records) {
    var priority = text(test.priority, '—');
    var hasCache = state.index.has(test.id);
    var stats = detailStatsHTML(records, hasCache);
    var formBusy = state.form && state.form.testId === test.id;
    var sourceChip = '<span class="source-chip">' + icon('info') + '原表参考</span>';
    return '<header class="detail-head">' +
      '<div class="detail-eyebrow">' +
      '<span>' + esc(test.id) + '</span><span class="sep">/</span>' +
      '<span class="module-name">' + esc(moduleNameOf(test)) + '</span>' +
      '<span class="badge prio-badge ' + priorityClass(priority) + '">' + esc(priority) + '</span>' +
      sourceChip +
      '</div>' +
      '<div class="detail-title-row">' +
      '<div>' +
      '<h2 class="detail-title">' + esc(text(test.title, '未命名测试项')) + '</h2>' +
      '<p class="detail-purpose">' + esc(text(test.purpose, '未填写测试目的')) + '</p>' +
      '</div>' +
      '<div class="detail-actions">' +
      '<button type="button" class="btn btn-quiet btn-sm" data-action="copy-test">' + icon('copy') + '复制</button>' +
      '<button type="button" class="btn btn-quiet btn-sm" data-action="cut-test">' + icon('cut') + '剪切</button>' +
      '<button type="button" class="btn btn-quiet" data-action="edit-guidance">' + icon('edit') + '编辑说明</button>' +
      '<button type="button" class="btn btn-quiet btn-danger-quiet" data-action="delete-test" title="删除该测试项及其记录">' + icon('trash') + '删除测试项</button>' +
      '<button type="button" class="btn btn-primary" data-action="new-record"' + (formBusy ? ' disabled title="请先完成当前记录"' : '') + '>' +
      icon('plus') + '新建记录</button>' +
      '</div>' +
      '</div>' +
      '<div class="detail-stats" id="detailStats">' + stats + '</div>' +
      '</header>';
  }

  function detailStatsHTML(records, hasCache) {
    if (!hasCache) {
      return detailStat('记录', '—') + detailStat('Pass', '—', 'pass') + detailStat('Fail', '—', 'fail') + detailStat('最近测试', '—');
    }
    var pass = 0;
    var fail = 0;
    var latestDate = '';
    records.forEach(function (record) {
      if (record.status === 'Pass') pass += 1;
      else if (record.status === 'Fail') fail += 1;
      if (record.test_date && record.test_date > latestDate) latestDate = record.test_date;
    });
    return detailStat('记录', String(records.length)) +
      detailStat('Pass', String(pass), 'pass') +
      detailStat('Fail', String(fail), 'fail') +
      detailStat('最近测试', latestDate || '—');
  }

  function detailStat(label, value, modifier) {
    return '<div class="detail-stat"><span>' + esc(label) + '</span><b class="' + (modifier || '') + '">' + esc(value) + '</b></div>';
  }

  function guidanceFields() {
    return [
      { key: 'stimulus', label: '测试连接/激励', cls: '' },
      { key: 'metrics', label: '需记录的指标', cls: '' },
      { key: 'conditions', label: '建议测试条件', cls: '' },
      { key: 'calculation', label: '计算/处理方法', cls: 'is-calculation', formula: true },
      { key: 'criteria', label: '论文参考/建议判据', cls: 'is-criteria' },
      { key: 'instruments', label: '仪器', cls: 'is-instrument' }
    ];
  }

  function guidanceHTML(test) {
    var parts = [];
    guidanceFields().forEach(function (field) {
      var value = text(test[field.key]);
      var valueHtml = value ? esc(value) : '<span class="source-empty">未填写</span>';
      parts.push('<article class="guidance-card ' + field.cls + '">' +
        '<div class="guidance-label">' + esc(field.label) + '</div>' +
        '<div class="guidance-value' + (field.formula && value ? ' is-formula' : '') + '">' + valueHtml + '</div>' +
        '</article>');
    });
    return '<section class="guidance-section" aria-label="测试指导">' +
      '<div class="guidance-head">' +
      '<h3 class="guidance-title">测试指导</h3>' +
      '<button type="button" class="text-btn" data-action="edit-guidance" title="编辑测试目的、激励、指标、条件、算法、判据与仪器">' +
      icon('edit') + '编辑说明</button>' +
      '</div>' +
      '<div class="guidance-grid">' + parts.join('') + '</div>' +
      '</section>';
  }

  function sourcePanelHTML(test) {
    var sourceFields = [
      { key: 'source_result', label: '原表结果' },
      { key: 'source_unit', label: '原表单位' },
      { key: 'source_status', label: '原表判定' },
      { key: 'source_date', label: '原表日期' },
      { key: 'source_notes', label: '原表备注' }
    ];
    var hasSource = false;
    sourceFields.forEach(function (field) { if (text(test[field.key])) hasSource = true; });

    var summaryRows = state.summary.filter(function (row) {
      return row && text(row['模块']) === moduleNameOf(test);
    });

    var sourceHtml = '<div class="source-grid">';
    sourceFields.forEach(function (field) {
      var value = text(test[field.key]);
      sourceHtml += '<dl class="source-field"><dt>' + esc(field.label) + '</dt><dd' + (value ? '' : ' class="is-empty"') + '>' +
        (value ? esc(value) : '—') + '</dd></dl>';
    });
    sourceHtml += '</div>';

    var summaryHtml = '';
    if (summaryRows.length) {
      summaryHtml = '<div class="summary-table-wrap"><table class="summary-table">' +
        '<thead><tr><th>建议优先记录的核心指标</th><th>原表结果</th><th>单位</th><th>目标 / 判据</th><th>备注</th></tr></thead><tbody>';
      summaryRows.forEach(function (row) {
        summaryHtml += '<tr>' +
          '<td>' + esc(text(row['建议优先记录的核心指标'], '—')) + '</td>' +
          '<td' + (text(row['结果']) ? '' : ' class="is-muted"') + '>' + esc(text(row['结果'], '—')) + '</td>' +
          '<td' + (text(row['单位']) ? '' : ' class="is-muted"') + '>' + esc(text(row['单位'], '—')) + '</td>' +
          '<td>' + esc(text(row['目标/判据'], '—')) + '</td>' +
          '<td class="is-muted">' + esc(text(row['备注'], '—')) + '</td>' +
          '</tr>';
      });
      summaryHtml += '</tbody></table></div>';
    } else {
      summaryHtml = '<p class="source-empty" style="margin:12px 0 0">该模块暂无原表核心指标建议。</p>';
    }

    var open = hasSource || summaryRows.length ? ' open' : '';
    return '<details class="source-panel"' + open + '>' +
      '<summary>' + icon('info') + '<span>原始表格参考</span><span class="source-tag">只读 · 不参与记录</span></summary>' +
      '<div class="source-body">' + sourceHtml + summaryHtml + '</div>' +
      '</details>';
  }

  function recordsAreaHTML(test) {
    var all = state.index.get(test.id);
    var hasCache = !!all;
    var list = hasCache ? all : [];
    var loading = state.recordsLoading.has(test.id);
    var error = state.recordErrors.get(test.id);
    var parts = [];

    parts.push('<div class="records-head">' +
      '<div>' +
      '<div class="records-title-row"><h3 class="records-title">记录历史</h3>' +
      '<span class="count-pill">' + (hasCache ? String(list.length) : '—') + '</span>' +
      (loading ? '<span class="spinner light" style="margin-left:4px" title="读取中"></span>' : '') +
      '</div>' +
      '<p class="records-sub">每次仿真 / 实测独立留档；先保存记录，再上传附件。</p>' +
      '</div>' +
      recordToolbarHTML(list, loading) +
      '</div>');

    if (!hasCache) {
      if (error) {
        parts.push('<div class="records-error">' + icon('alert') +
          '<h4>记录读取失败</h4><p>' + esc(error) + '</p>' +
          '<button type="button" class="btn btn-sm" data-action="retry-records">' + icon('refresh') + '重试读取</button>' +
          '</div>');
      } else {
        parts.push('<div class="records-loading"><div class="skeleton skeleton-line"></div><div class="skeleton skeleton-line"></div></div>');
      }
      return parts.join('');
    }

    if (error) {
      parts.push('<div class="alert alert-warn" style="margin-top:11px">' +
        '<div class="alert-icon">' + icon('alert') + '</div>' +
        '<div class="alert-body"><p class="alert-title">记录刷新失败，当前显示上次读取的数据</p>' +
        '<p class="alert-text">' + esc(error) + '</p>' +
        '<div class="alert-actions"><button type="button" class="btn btn-sm" data-action="retry-records">' + icon('refresh') + '重试刷新</button></div>' +
        '</div></div>');
    }

    var filtered = filterRecordList(list);
    if (!list.length) {
      parts.push('<div class="empty-records">' +
        '<div class="empty-icon">' + icon('database') + '</div>' +
        '<h4>尚无测试记录</h4>' +
        '<p>完成仿真或实测后，在此登记结果、判定、日期和备注，可追加 Excel 数据与 JPG 图片。</p>' +
        '<button type="button" class="btn btn-primary btn-sm" data-action="new-record">' + icon('plus') + '新建第一条记录</button>' +
        '</div>');
    } else if (!filtered.length) {
      parts.push('<div class="empty-records">' +
        '<div class="empty-icon">' + icon('search') + '</div>' +
        '<h4>没有符合筛选条件的记录</h4>' +
        '<p>当前状态或方式筛选未匹配到记录，可清除筛选查看全部。</p>' +
        '<button type="button" class="btn btn-sm" data-action="clear-record-filters">清除记录筛选</button>' +
        '</div>');
    } else {
      var cards = [];
      var latestId = list.length ? list[list.length - 1].id : 0;
      filtered.forEach(function (record) {
        cards.push(recordCardHTML(record, record.id === latestId, test));
      });
      parts.push('<div class="record-list">' + cards.join('') + '</div>');
    }
    return parts.join('');
  }

  function recordToolbarHTML(list, loading) {
    var counts = { all: list.length, Pass: 0, Fail: 0, blank: 0, simulation: 0, measurement: 0 };
    list.forEach(function (record) {
      if (record.status === 'Pass') counts.Pass += 1;
      else if (record.status === 'Fail') counts.Fail += 1;
      else counts.blank += 1;
      if (record.mode === 'simulation') counts.simulation += 1;
      if (record.mode === 'measurement') counts.measurement += 1;
    });
    var disabled = list.length ? '' : ' disabled';
    var refreshDisabled = loading ? ' disabled' : '';
    var filters = state.recordFilters;
    return '<div class="record-tools">' +
      '<span class="toolbar-label">判定</span>' +
      '<div class="segmented" role="group" aria-label="按判定筛选记录">' +
      recordFilterButton('all', '全部', counts.all, filters.status === 'all', disabled, '') +
      recordFilterButton('Pass', 'Pass', counts.Pass, filters.status === 'Pass', disabled, 'pass') +
      recordFilterButton('Fail', 'Fail', counts.Fail, filters.status === 'Fail', disabled, 'fail') +
      recordFilterButton('blank', '未判定', counts.blank, filters.status === 'blank', disabled, '') +
      '</div>' +
      '<span class="toolbar-label">方式</span>' +
      '<div class="segmented" role="group" aria-label="按方式筛选记录">' +
      recordModeButton('all', '全部', counts.all, filters.mode === 'all', disabled) +
      recordModeButton('simulation', '仿真', counts.simulation, filters.mode === 'simulation', disabled) +
      recordModeButton('measurement', '实测', counts.measurement, filters.mode === 'measurement', disabled) +
      '</div>' +
      '<button type="button" class="icon-btn" data-action="refresh-records" title="刷新记录"' + refreshDisabled + ' aria-label="刷新记录">' +
      icon('refresh', loading ? 'is-spinning' : '') + '</button>' +
      '</div>';
  }

  function recordFilterButton(value, label, count, active, disabled, extra) {
    return '<button type="button" class="seg-btn' + (active ? ' active ' + extra : '') + '" data-record-status="' + esc(value) + '"' +
      ' aria-pressed="' + (active ? 'true' : 'false') + '"' + disabled + '>' + esc(label) + ' ' + count + '</button>';
  }

  function recordModeButton(value, label, count, active, disabled) {
    return '<button type="button" class="seg-btn' + (active ? ' active' : '') + '" data-record-mode="' + esc(value) + '"' +
      ' aria-pressed="' + (active ? 'true' : 'false') + '"' + disabled + '>' + esc(label) + ' ' + count + '</button>';
  }

  function recordMatchesFilters(record) {
    if (state.recordFilters.status !== 'all') {
      if (state.recordFilters.status === 'blank' && record.status !== '') return false;
      if (state.recordFilters.status !== 'blank' && record.status !== state.recordFilters.status) return false;
    }
    if (state.recordFilters.mode !== 'all' && record.mode !== state.recordFilters.mode) return false;
    return true;
  }

  function filterRecordList(records) {
    return records.filter(recordMatchesFilters);
  }

  function recordCardHTML(record, isLatest, test) {
    var dotClass = record.status === 'Pass' ? 'pass' : (record.status === 'Fail' ? 'fail' : 'blank');
    var resultHtml = record.result
      ? '<span class="record-result">' + esc(record.result) + (record.unit ? '<em class="unit">' + esc(record.unit) + '</em>' : '') + '</span>'
      : (record.unit
        ? '<span class="record-result">—<em class="unit">' + esc(record.unit) + '</em></span>'
        : '<span class="record-result no-value">未填写结果</span>');

    return '<article class="record-card' + (isLatest ? ' is-latest' : '') + '">' +
      '<div class="record-rail"><span class="dot ' + dotClass + '"></span></div>' +
      '<div class="record-main">' +
      '<div class="record-head">' +
      '<span class="chip mode-chip ' + (record.mode || 'unknown') + '">' + esc(modeLabel(record.mode)) + '</span>' +
      '<span class="chip status-chip ' + (record.status === 'Pass' ? 'pass' : (record.status === 'Fail' ? 'fail' : 'blank')) + '">' + esc(statusLabel(record.status)) + '</span>' +
      '<span class="record-date">' + icon('clock') + esc(record.test_date || '未填写日期') + '</span>' +
      resultHtml +
      '<span class="record-actions">' +
      '<button type="button" class="icon-btn" data-action="edit-record" data-record-id="' + record.id + '" title="编辑记录" aria-label="编辑记录"' + (state.submitting ? ' disabled' : '') + '>' +
      icon('edit') + '</button>' +
      '</span>' +
      '</div>' +
      '<div class="record-meta">记录 #' + record.id + ' · 创建 ' + esc(formatDateTime(record.created_at)) +
      (record.updated_at && record.updated_at !== record.created_at ? ' · 更新 ' + esc(formatDateTime(record.updated_at)) : '') +
      '</div>' +
      (record.notes ? '<p class="record-notes">' + esc(record.notes) + '</p>' : '') +
      attachmentsHTML(record) +
      '</div>' +
      '</article>';
  }

  function attachmentsHTML(record) {
    var attachments = record.attachments || [];
    if (!attachments.length) return '';
    var data = [];
    var images = [];
    attachments.forEach(function (att) {
      if (att.kind === 'image') images.push(att);
      else data.push(att);
    });
    var parts = [];
    if (data.length) {
      data.forEach(function (att) {
        var url = attachmentUrl(att);
        parts.push('<div class="file-chip">' +
          '<a class="file-chip-link" href="' + esc(url) + '" download="' + esc(att.name) + '" title="下载 ' + esc(att.name) + '">' +
          '<span class="file-glyph">' + icon('file') + '</span>' +
          '<span class="file-chip-name">' + esc(att.name) + '</span>' +
          '<span class="file-chip-size">' + esc(formatBytes(att.size)) + '</span>' +
          '</a>' +
          removeAttachmentButton(record, att, '移除数据附件') +
          '</div>');
      });
    }
    var thumbs = '';
    if (images.length) {
      var cells = [];
      images.forEach(function (att) {
        var url = safeUrl(attachmentUrl(att));
        cells.push('<div class="thumb">' +
          '<button type="button" class="thumb-open" data-action="open-image" data-url="' + esc(url) + '" data-name="' + esc(att.name) + '" title="预览 ' + esc(att.name) + '">' +
          '<img src="' + esc(url) + '" alt="' + esc(att.name) + '" loading="lazy">' +
          '</button>' +
          '<span class="thumb-name" title="' + esc(att.name) + '">' + esc(att.name) + '</span>' +
          '<button type="button" class="thumb-remove" data-action="remove-attachment" data-attachment-id="' + att.id + '" data-record-id="' + record.id + '" data-attachment-name="' + esc(att.name) + '" title="移除图片" aria-label="移除图片 ' + esc(att.name) + '">' +
          icon('close') + '</button>' +
          '</div>');
      });
      thumbs = '<div class="thumb-grid">' + cells.join('') + '</div>';
    }
    return '<div class="attachments">' + parts.join('') + thumbs + '</div>';
  }

  function removeAttachmentButton(record, att, label) {
    return '<button type="button" class="remove-att" data-action="remove-attachment" data-attachment-id="' + att.id + '"' +
      ' data-record-id="' + record.id + '" data-attachment-name="' + esc(att.name) + '" title="' + esc(label) + '" aria-label="' + esc(label) + ' ' + esc(att.name) + '">' +
      icon('close') + '</button>';
  }

  function wireImageFallbacks() {
    $$('.thumb-open img').forEach(function (img) {
      if (img.getAttribute('data-wired') === '1') return;
      img.setAttribute('data-wired', '1');
      img.addEventListener('error', function () {
        img.classList.add('is-broken');
        img.alt = '图片加载失败';
        img.title = '图片加载失败，可尝试刷新记录';
      });
    });
    $$('#lightboxImage').forEach(function (img) {
      if (img.getAttribute('data-wired') === '1') return;
      img.setAttribute('data-wired', '1');
      img.addEventListener('error', function () {
        var caption = $('#lightboxCaption');
        if (caption) caption.textContent = '图片加载失败';
      });
    });
  }

  function renderDetailStats() {
    var test = selectedTest();
    var host = $('#detailStats');
    if (!test || !host) return;
    var records = state.index.get(test.id) || [];
    host.innerHTML = detailStatsHTML(records, state.index.has(test.id));
  }

  function renderRecordsArea() {
    var test = selectedTest();
    var host = $('#recordsArea');
    if (!test || !host) return;
    host.innerHTML = recordsAreaHTML(test);
    wireImageFallbacks();
  }

  function renderDetailAlert() {
    var host = $('#detailAlert');
    if (!host) return;
    host.innerHTML = detailAlertHTML();
  }

  function detailAlertHTML() {
    var pending = state.pendingUpload;
    var test = selectedTest();
    if (!pending || !test || pending.testId !== test.id) return '';
    var failedCount = 0;
    pending.groups.forEach(function (group) { failedCount += group.files.length; });
    var title = pending.busy ? '正在上传附件…' : (pending.fromSave ? '记录已保存，附件上传失败' : '附件重试上传失败');
    var textHtml = '记录 <strong>#' + pending.recordId + '</strong> 已保存。';
    if (pending.uploaded) textHtml += ' 已成功上传 <strong>' + pending.uploaded + '</strong> 个附件。';
    if (pending.busy) {
      textHtml += ' 正在上传 <strong>' + failedCount + '</strong> 个附件，请勿离开或刷新页面。';
    } else {
      textHtml += ' 仍有 <strong>' + failedCount + '</strong> 个附件未上传，可重试；重试不会重复创建记录。';
    }
    var errors = [];
    pending.groups.forEach(function (group) {
      var label = group.kind === 'image' ? '图片附件' : '数据附件';
      var names = group.files.map(function (file) { return file.name; });
      var shown = names.slice(0, 3).join('、');
      if (names.length > 3) shown += ' 等 ' + names.length + ' 个文件';
      errors.push(label + '（' + shown + '）：' + group.error);
    });
    var actions = '';
    if (pending.busy) {
      actions = '<div class="alert-actions"><button type="button" class="btn btn-sm" disabled><span class="spinner light"></span>正在上传…</button></div>';
    } else {
      actions = '<div class="alert-actions">' +
        '<button type="button" class="btn btn-primary btn-sm" data-action="retry-upload">' + icon('upload') + '重试上传</button>' +
        '<button type="button" class="text-btn" data-action="discard-upload">放弃附件</button>' +
        '</div>';
    }
    return '<div class="alert alert-warn" role="alert">' +
      '<div class="alert-icon">' + icon('alert') + '</div>' +
      '<div class="alert-body">' +
      '<p class="alert-title">' + esc(title) + '</p>' +
      '<p class="alert-text">' + textHtml + '</p>' +
      (errors.length ? '<p class="alert-text">' + esc(errors.join('；')) + '</p>' : '') +
      actions +
      '</div></div>';
  }

  /* ------------------------------------------------------------------------
     record form
     ------------------------------------------------------------------------ */

  function newFormValues() {
    return { mode: '', result: '', unit: '', status: '', test_date: todayISO(), notes: '' };
  }

  async function confirmDiscardPendingForForm() {
    var pending = state.pendingUpload;
    if (!pending) return true;
    if (pending.busy) {
      showToast('附件正在上传，请稍候。', 'info');
      return false;
    }
    var count = 0;
    pending.groups.forEach(function (group) { count += group.files.length; });
    var ok = await confirmDialog({
      title: '有附件待上传',
      message: '记录 #' + pending.recordId + ' 还有 ' + count + ' 个附件未上传。继续填写新记录将放弃这些待上传文件，已保存的记录不受影响。',
      confirmText: '放弃并继续',
      cancelText: '返回处理',
      danger: true
    });
    if (!ok) return false;
    state.pendingUpload = null;
    renderDetailAlert();
    return true;
  }

  async function openNewForm() {
    if (state.phase !== 'ready' || !state.selectedId) return;
    if (state.submitting || (state.pendingUpload && state.pendingUpload.busy)) {
      showToast('请等待当前操作完成。', 'info');
      return;
    }
    if (state.form) {
      if (state.form.kind === 'new') {
        scrollToForm();
      } else {
        showToast('请先完成或取消当前记录的编辑。', 'info');
      }
      return;
    }
    if (!(await confirmDiscardPendingForForm())) return;
    var test = selectedTest();
    if (!test) return;
    var values = newFormValues();
    state.form = {
      kind: 'new',
      testId: test.id,
      recordId: null,
      values: values,
      initial: JSON.stringify(values),
      dataFiles: [],
      imageFiles: [],
      previews: new Map(),
      errors: {},
      alert: null,
      submitting: false
    };
    renderDetail();
    focusFormSoon();
  }

  async function openEditForm(recordId) {
    if (!recordId || state.submitting) return;
    if (state.pendingUpload && state.pendingUpload.busy) {
      showToast('附件正在上传，请稍候。', 'info');
      return;
    }
    var test = selectedTest();
    if (!test) return;
    var record = findRecord(test.id, recordId);
    if (!record) {
      showToast('未找到该记录，可能已被移除。', 'error');
      return;
    }
    if (state.form && state.form.kind === 'edit' && state.form.recordId === recordId) {
      scrollToForm();
      return;
    }
    var ok = await guardUnsavedForm('编辑其他记录');
    if (!ok) return;
    if (!(await confirmDiscardPendingForForm())) return;

    var values = {
      mode: record.mode || '',
      result: record.result || '',
      unit: record.unit || '',
      status: record.status || '',
      test_date: record.test_date || todayISO(),
      notes: record.notes || ''
    };
    state.form = {
      kind: 'edit',
      testId: test.id,
      recordId: record.id,
      values: values,
      initial: JSON.stringify(values),
      dataFiles: [],
      imageFiles: [],
      previews: new Map(),
      errors: {},
      alert: null,
      submitting: false
    };
    renderDetail();
    focusFormSoon();
  }

  function focusFormSoon() {
    requestAnimationFrame(function () {
      var form = $('#recordForm');
      if (!form) return;
      form.scrollIntoView({ block: 'start', behavior: 'smooth' });
      var first = form.querySelector('input[name="record-mode"]');
      if (first) first.focus({ preventScroll: true });
    });
  }

  function scrollToForm() {
    var form = $('#recordForm');
    if (!form) return;
    form.scrollIntoView({ block: 'start', behavior: 'smooth' });
  }

  function snapshotForm() { return state.form ? JSON.stringify(state.form.values) : ''; }

  function isFormDirty() {
    var form = state.form;
    if (!form) return false;
    if (form.dataFiles.length || form.imageFiles.length) return true;
    return JSON.stringify(form.values) !== form.initial;
  }

  async function guardUnsavedForm(actionLabel) {
    if (!state.form) return true;
    if (state.submitting || state.form.submitting) {
      showToast('正在保存记录，请稍候。', 'info');
      return false;
    }
    var dirty = isFormDirty();
    if (!dirty) {
      closeFormInternal();
      return true;
    }
    var ok = await confirmDialog({
      title: '放弃未保存的更改？',
      message: '当前记录尚未保存。' + (actionLabel ? actionLabel + '后，' : '离开后，') + '已填写的内容和已选择的附件都会丢失。',
      note: '附件只有在上传成功后才会写入服务器。',
      confirmText: '放弃更改',
      cancelText: '继续编辑',
      danger: true
    });
    if (!ok) return false;
    closeFormInternal();
    return true;
  }

  async function guardNavigation() {
    if (state.submitting || (state.form && state.form.submitting)) {
      showToast('正在保存记录，请稍候。', 'info');
      return false;
    }
    if (state.pendingUpload && state.pendingUpload.busy) {
      showToast('附件正在上传，请稍候。', 'info');
      return false;
    }
    var pending = state.pendingUpload;
    if (pending) {
      var count = 0;
      pending.groups.forEach(function (group) { count += group.files.length; });
      var leave = await confirmDialog({
        title: '附件尚未上传',
        message: '记录 #' + pending.recordId + ' 已保存，但还有 ' + count + ' 个附件待上传。离开后这些待上传文件会丢失，已保存的记录不受影响。',
        confirmText: '仍然离开',
        cancelText: '留在此页',
        danger: true
      });
      if (!leave) return false;
    }
    if (!(await guardUnsavedForm('切换页面'))) return false;
    if (pending) state.pendingUpload = null;
    return true;
  }

  function closeFormInternal() {
    if (!state.form) return;
    state.form.previews.forEach(function (url) { URL.revokeObjectURL(url); });
    state.form.previews.clear();
    state.form = null;
  }

  async function cancelForm() {
    if (state.submitting || (state.form && state.form.submitting)) {
      showToast('正在保存记录，请稍候。', 'info');
      return;
    }
    var ok = await guardUnsavedForm('取消编辑');
    if (!ok) return;
    renderDetail();
  }

  function formHTML() {
    var form = state.form;
    if (!form) return '';
    var v = form.values;
    var dis = form.submitting ? ' disabled' : '';
    var blocked = !!form.blockedByAmbiguity;
    var modeError = form.errors.mode;
    var statusError = form.errors.status;
    var dateError = form.errors.test_date;
    var resultError = form.errors.result;
    var unitError = form.errors.unit;
    var notesError = form.errors.notes;

    var parts = [];
    parts.push('<form class="record-form' + (form.submitting ? ' is-submitting' : '') + '" id="recordForm" novalidate aria-busy="' + (form.submitting ? 'true' : 'false') + '">');
    parts.push('<div class="form-head">' +
      '<div><h3 class="form-title">' + (form.kind === 'edit' ? '编辑记录 #' + form.recordId : '新建测试记录') + '</h3>' +
      '<p class="form-sub">' + (form.kind === 'edit'
        ? '保存修改后，可继续为这条记录添加 Excel 或 JPG 附件。已有附件可在记录卡片中移除。'
        : '填写本次仿真或实测结果。先保存记录，再上传所选附件；附件失败可单独重试。') +
      '</p></div>' +
      '<button type="button" class="icon-btn" data-action="cancel-form" aria-label="关闭表单" title="关闭表单"' + dis + '>' + icon('close') + '</button>' +
      '</div>');
    if (form.alert) parts.push(formAlertHTML(form.alert));
    parts.push('<div class="form-grid">');

    parts.push('<div class="field' + (modeError ? ' is-invalid' : '') + '" data-field-wrap="mode" role="radiogroup" aria-labelledby="modeLabel">' +
      '<div class="field-label" id="modeLabel">记录方式 <span class="req">*</span><span class="label-hint">必填</span></div>' +
      '<div class="radio-row">' +
      '<label class="radio-pill is-sim"><input type="radio" name="record-mode" value="simulation" data-field="mode"' + (v.mode === 'simulation' ? ' checked' : '') + dis + '>' +
      '<span class="radio-body">仿真 <small>Simulation</small></span></label>' +
      '<label class="radio-pill is-meas"><input type="radio" name="record-mode" value="measurement" data-field="mode"' + (v.mode === 'measurement' ? ' checked' : '') + dis + '>' +
      '<span class="radio-body">实测 <small>Measurement</small></span></label>' +
      '</div>' +
      (modeError ? '<p class="field-error" data-field-error="mode">' + esc(modeError) + '</p>' : '') +
      '</div>');

    parts.push('<div class="field' + (statusError ? ' is-invalid' : '') + '" data-field-wrap="status" role="radiogroup" aria-labelledby="statusLabel">' +
      '<div class="field-label" id="statusLabel">判定结果 <span class="req">*</span><span class="label-hint">最终结论</span></div>' +
      '<div class="radio-row">' +
      '<label class="radio-pill is-pass"><input type="radio" name="record-status" value="Pass" data-field="status"' + (v.status === 'Pass' ? ' checked' : '') + dis + '>' +
      '<span class="radio-body">Pass</span></label>' +
      '<label class="radio-pill is-fail"><input type="radio" name="record-status" value="Fail" data-field="status"' + (v.status === 'Fail' ? ' checked' : '') + dis + '>' +
      '<span class="radio-body">Fail</span></label>' +
      '</div>' +
      (statusError ? '<p class="field-error" data-field-error="status">' + esc(statusError) + '</p>' : '') +
      '</div>');

    parts.push('<label class="field' + (dateError ? ' is-invalid' : '') + '" data-field-wrap="test_date">' +
      '<span class="field-label">测试日期 <span class="req">*</span><span class="label-hint">YYYY-MM-DD</span></span>' +
      '<input type="date" name="test_date" data-field="test_date" value="' + esc(v.test_date) + '"' + dis + '>' +
      (dateError ? '<p class="field-error" data-field-error="test_date">' + esc(dateError) + '</p>' : '') +
      '</label>');

    parts.push('<label class="field' + (resultError ? ' is-invalid' : '') + '" data-field-wrap="result">' +
      '<span class="field-label">结果 <span class="label-hint">可选 · 最多 500 字</span></span>' +
      '<input type="text" name="result" data-field="result" value="' + esc(v.result) + '" maxlength="500" placeholder="如 1.253 或 &lt; 0.5"' + dis + '>' +
      (resultError ? '<p class="field-error" data-field-error="result">' + esc(resultError) + '</p>' : '') +
      '</label>');

    parts.push('<label class="field' + (unitError ? ' is-invalid' : '') + '" data-field-wrap="unit">' +
      '<span class="field-label">单位 <span class="label-hint">可选 · 最多 100 字</span></span>' +
      '<input type="text" name="unit" data-field="unit" value="' + esc(v.unit) + '" maxlength="100" placeholder="如 kΩ / dB / ns"' + dis + '>' +
      (unitError ? '<p class="field-error" data-field-error="unit">' + esc(unitError) + '</p>' : '') +
      '</label>');

    parts.push('<label class="field span-2' + (notesError ? ' is-invalid' : '') + '" data-field-wrap="notes">' +
      '<span class="field-label">备注 <span class="label-hint">可选 · 测试条件、仪器设置、异常说明</span></span>' +
      '<textarea name="notes" data-field="notes" rows="3" maxlength="20000" placeholder="记录测试条件、仪器带宽、采样率、异常现象或后续动作…"' + dis + '>' + esc(v.notes) + '</textarea>' +
      (notesError ? '<p class="field-error" data-field-error="notes">' + esc(notesError) + '</p>' : '') +
      '</label>');

    parts.push(uploadFieldHTML('data', form, dis));
    parts.push(uploadFieldHTML('image', form, dis));

    parts.push('</div>');
    parts.push('<div class="form-footer">' +
      '<span class="form-hint">' + icon('info') + '先保存记录，再上传附件；上传失败可重试且不会重复创建记录。</span>' +
      '<div class="form-buttons">' +
      '<button type="button" class="btn" data-action="cancel-form"' + dis + '>取消</button>' +
      '<button type="submit" class="btn btn-primary"' + (form.submitting || blocked ? ' disabled' : '') + '>' +
      (form.submitting ? '<span class="spinner"></span>正在保存记录…' : (blocked ? '请先刷新记录确认' : submitLabel(form))) +
      '</button>' +
      '</div></div>');
    parts.push('</form>');
    return parts.join('');
  }

  function submitLabel(form) {
    var hasFiles = form.dataFiles.length + form.imageFiles.length > 0;
    if (form.kind === 'edit') return hasFiles ? '保存修改并上传附件' : '保存修改';
    return hasFiles ? '保存记录并上传附件' : '保存记录';
  }

  function formAlertHTML(alert) {
    var type = alert.type === 'error' ? 'alert-error' : 'alert-warn';
    var iconName = alert.type === 'error' ? 'alert' : 'info';
    var parts = ['<div class="alert ' + type + ' form-alert" role="alert">' +
      '<div class="alert-icon">' + icon(iconName) + '</div>' +
      '<div class="alert-body">' +
      '<p class="alert-title">' + esc(alert.title) + '</p>' +
      '<p class="alert-text">' + esc(alert.text) + '</p>'];
    if (alert.showRefresh) {
      parts.push('<div class="alert-actions"><button type="button" class="btn btn-sm" data-action="refresh-from-form">' + icon('refresh') + '刷新记录确认</button></div>');
    }
    parts.push('</div></div>');
    return parts.join('');
  }

  function uploadFieldHTML(kind, form, dis) {
    var isData = kind === 'data';
    var files = isData ? form.dataFiles : form.imageFiles;
    var label = isData ? '数据附件' : '图片附件';
    var hint = isData ? 'Excel · 可多选' : 'JPG · 可多选 · 自动预览';
    var accept = isData ? '.xlsx,.xls,.xlsm' : '.jpg,.jpeg,image/jpeg';
    var dropId = isData ? 'dataDropzone' : 'imageDropzone';
    var inputId = isData ? 'dataFileInput' : 'imageFileInput';
    var listId = isData ? 'dataFileList' : 'imageFileGrid';
    var listMarkup = isData ? dataFilesHTML(form) : imageFilesHTML(form);
    return '<div class="field span-2 upload-field" data-field-wrap="' + kind + '_files">' +
      '<div class="upload-head"><span class="field-label">' + label + ' <span class="label-hint">' + hint + '</span></span>' +
      '<small>' + (isData ? 'xlsx / xls / xlsm · 单个 ≤ 50 MiB' : 'jpg / jpeg · 单个 ≤ 50 MiB') + '</small></div>' +
      '<div class="dropzone" id="' + dropId + '" data-dropzone="' + kind + '" role="button" tabindex="' + (dis ? '-1' : '0') + '" aria-label="选择' + label + '">' +
      '<input type="file" id="' + inputId + '" data-file-kind="' + kind + '" accept="' + accept + '" multiple tabindex="-1"' + dis + '>' +
      '<span class="dropzone-icon">' + icon(isData ? 'file' : 'image') + '</span>' +
      '<span class="dropzone-copy"><span class="dropzone-title">' + (isData ? '点击或拖拽 Excel 文件到此' : '点击或拖拽 JPG 图片到此') + '</span>' +
      '<span class="dropzone-hint">' + (isData ? '支持 xlsx / xls / xlsm，单个文件不超过 50 MiB' : '仅支持 jpg / jpeg，添加后可预览，单个不超过 50 MiB') + '</span></span>' +
      '</div>' +
      '<div class="' + (isData ? 'file-list' : 'image-file-grid') + '" id="' + listId + '">' + listMarkup + '</div>' +
      '</div>';
  }

  function dataFilesHTML(form) {
    if (!form.dataFiles.length) return '';
    var parts = [];
    form.dataFiles.forEach(function (file, index) {
      parts.push('<div class="file-row">' +
        '<span class="file-glyph">' + icon('file') + '</span>' +
        '<span class="file-info"><span class="file-name" title="' + esc(file.name) + '">' + esc(file.name) + '</span>' +
        '<span class="file-size">' + esc(formatBytes(file.size)) + '</span></span>' +
        '<button type="button" class="file-remove" data-action="remove-data-file" data-index="' + index + '" title="移除文件" aria-label="移除 ' + esc(file.name) + '">' + icon('close') + '</button>' +
        '</div>');
    });
    return parts.join('');
  }

  function imageFilesHTML(form) {
    if (!form.imageFiles.length) return '';
    var parts = [];
    form.imageFiles.forEach(function (file, index) {
      var url = form.previews.get(file) || '';
      parts.push('<div class="image-file">' +
        '<img src="' + esc(url) + '" alt="预览 ' + esc(file.name) + '">' +
        '<span class="image-file-body"><span class="image-file-name" title="' + esc(file.name) + '">' + esc(file.name) + '</span>' +
        '<span class="image-file-size">' + esc(formatBytes(file.size)) + '</span></span>' +
        '<button type="button" class="image-file-remove" data-action="remove-image-file" data-index="' + index + '" title="移除图片" aria-label="移除 ' + esc(file.name) + '">' + icon('close') + '</button>' +
        '</div>');
    });
    return parts.join('');
  }

  function renderForm() {
    var host = $('#formHost');
    if (!host) return;
    host.innerHTML = state.form ? formHTML() : '';
    wireImageFallbacks();
  }

  function clearFieldError(name) {
    var form = state.form;
    if (!form) return;
    delete form.errors[name];
    var wrap = document.querySelector('[data-field-wrap="' + name + '"]');
    if (!wrap) return;
    wrap.classList.remove('is-invalid');
    var err = wrap.querySelector('[data-field-error]');
    if (err) err.remove();
  }

  function setFormValue(input) {
    var form = state.form;
    if (!form || form.submitting) return;
    var name = input.getAttribute('data-field');
    if (!name) return;
    if (input.type === 'radio' && !input.checked) return;
    form.values[name] = input.value;
    clearFieldError(name);
  }

  function validateForm(form) {
    var errors = {};
    var v = form.values;
    if (v.mode !== 'simulation' && v.mode !== 'measurement') errors.mode = '请选择记录方式（仿真或实测）。';
    if (v.status !== 'Pass' && v.status !== 'Fail') errors.status = '请选择判定结果（Pass 或 Fail）。';
    if (!v.test_date) errors.test_date = '请选择测试日期。';
    else if (!isValidISODate(v.test_date)) errors.test_date = '日期格式无效，请重新选择。';
    if (v.result.length > 500) errors.result = '结果不能超过 500 字。';
    if (v.unit.length > 100) errors.unit = '单位不能超过 100 字。';
    if (v.notes.length > 20000) errors.notes = '备注不能超过 20000 字。';
    return errors;
  }

  function focusFirstError(errors) {
    var order = ['mode', 'status', 'test_date', 'result', 'unit', 'notes'];
    for (var i = 0; i < order.length; i++) {
      if (!errors[order[i]]) continue;
      var wrap = document.querySelector('[data-field-wrap="' + order[i] + '"]');
      if (!wrap) continue;
      var input = wrap.querySelector('input:not([type="file"]), textarea');
      if (input) {
        input.focus({ preventScroll: true });
        wrap.scrollIntoView({ block: 'center', behavior: 'smooth' });
        return;
      }
    }
  }

  async function submitForm() {
    var form = state.form;
    if (!form || state.submitting || form.submitting) return;
    if (state.pendingUpload) {
      form.alert = {
        type: 'error',
        title: '存在待上传附件',
        text: '记录 #' + state.pendingUpload.recordId + ' 还有附件待上传。请先在页面顶部重试上传或放弃附件，再保存当前记录。'
      };
      renderForm();
      showToast('请先处理待上传附件。', 'warning');
      return;
    }
    if (form.blockedByAmbiguity) {
      showToast('上次保存结果未知，请先点击“刷新记录确认”。', 'warning');
      return;
    }
    var errors = validateForm(form);
    if (Object.keys(errors).length) {
      form.errors = errors;
      form.alert = null;
      renderForm();
      focusFirstError(errors);
      showToast('请补全必填信息后再保存。', 'warning');
      return;
    }

    var dataFiles = form.dataFiles.slice();
    var imageFiles = form.imageFiles.slice();
    var testId = form.testId;
    var recordId = form.recordId;
    var isEdit = form.kind === 'edit';
    var payload = {
      mode: form.values.mode,
      result: form.values.result.trim(),
      unit: form.values.unit.trim(),
      status: form.values.status,
      test_date: form.values.test_date,
      notes: form.values.notes.trim()
    };

    state.submitting = true;
    form.submitting = true;
    form.errors = {};
    form.alert = null;
    renderForm();

    var path = isEdit
      ? '/api/records/' + encodeURIComponent(recordId)
      : '/api/tests/' + encodeURIComponent(testId) + '/records';
    var method = isEdit ? 'PUT' : 'POST';
    var saved = false;

    try {
      var data = await apiJson(path, { method: method, json: payload, timeout: 45000 });
      if (!data || !data.record) {
        throw new RequestError('服务器未返回保存后的记录。', 0, 'bad-response');
      }
      var record = normalizeRecord(data.record);
      saved = true;
      state.submitting = false;
      if (!recordMatchesFilters(record)) {
        state.recordFilters = Object.assign({}, RECORD_FILTER_DEFAULT);
      }
      mergeRecord(testId, record);
      closeFormInternal();
      renderDetail();
      renderIndexDerived();

      var groups = [];
      if (dataFiles.length) groups.push({ kind: 'data', files: dataFiles });
      if (imageFiles.length) groups.push({ kind: 'image', files: imageFiles });

      if (!groups.length) {
        showToast(isEdit ? '记录已更新。' : '记录已保存。', 'success');
        try {
          await refreshRecords(testId, { force: true, notify: false });
        } catch (err) {
          showToast('记录已保存，但记录列表刷新失败，请点击刷新重试。', 'warning', 6500);
        }
      } else {
        showToast(isEdit ? '记录已更新，正在上传附件…' : '记录已保存，正在上传附件…', 'info', 3500);
        await startPendingUpload(record.id, testId, groups, 0, true);
      }
    } catch (err) {
      state.submitting = false;
      if (saved) {
        showToast('记录已保存，但后续界面更新失败：' + errorMessage(err) + '。请刷新记录查看。', 'error', 8000);
        try { renderDetail(); } catch (renderErr) { /* keep current view */ }
        return;
      }
      if (state.form === form) {
        form.submitting = false;
        var base = '记录未保存：' + errorMessage(err);
        if (ambiguousError(err)) {
          base += ' 保存请求的结果未知，服务器可能已写入记录。请先点击“刷新记录确认”，查看历史中是否已出现该记录，避免重复提交。';
        }
        form.alert = {
          type: 'error',
          title: '记录保存失败',
          text: base,
          showRefresh: ambiguousError(err)
        };
        if (ambiguousError(err)) form.blockedByAmbiguity = true;
        renderForm();
      }
      showToast('记录保存失败：' + errorMessage(err), 'error', 7000);
    }
  }

  /* ------------------------------------------------------------------------
     uploads and attachments
     ------------------------------------------------------------------------ */

  async function startPendingUpload(recordId, testId, groups, uploaded, fromSave) {
    var pending = {
      recordId: recordId,
      testId: testId,
      groups: groups.map(function (group) {
        return { kind: group.kind, files: group.files.slice(), error: '' };
      }),
      busy: true,
      uploaded: uploaded || 0,
      fromSave: !!fromSave
    };
    state.pendingUpload = pending;
    renderDetailAlert();
    renderRecordsArea();

    var added = 0;
    var failed = [];
    for (var i = 0; i < pending.groups.length; i++) {
      var group = pending.groups[i];
      try {
        var attachments = await uploadFiles(recordId, group);
        mergeAttachments(testId, recordId, attachments);
        added += attachments.length;
      } catch (err) {
        failed.push({ kind: group.kind, files: group.files, error: errorMessage(err) });
      }
    }

    if (state.pendingUpload !== pending) return;
    pending.uploaded += added;
    pending.busy = false;

    var failedCount = 0;
    failed.forEach(function (group) { failedCount += group.files.length; });

    if (failed.length) {
      pending.groups = failed;
      state.pendingUpload = pending;
      renderDetailAlert();
      renderRecordsArea();
      showToast((fromSave ? '记录已保存，但 ' : '重试后仍有 ') + failedCount + ' 个附件上传失败，可在提示条中重试。', 'error', 7000);
    } else {
      state.pendingUpload = null;
      renderDetailAlert();
      showToast(fromSave ? '记录已保存，' + added + ' 个附件已上传。' : '附件已上传（' + added + ' 个）。', 'success');
    }

    try {
      await refreshRecords(testId, { force: true, notify: false });
    } catch (err) {
      showToast('附件处理完成，但记录列表刷新失败；请点击刷新重试。', 'warning', 6500);
    }
    if (state.pendingUpload === pending) renderDetailAlert();
    renderIndexDerived();
  }

  function retryPendingUpload() {
    var pending = state.pendingUpload;
    if (!pending || pending.busy) return;
    var groups = pending.groups.map(function (group) {
      return { kind: group.kind, files: group.files.slice() };
    });
    if (!groups.length) return;
    startPendingUpload(pending.recordId, pending.testId, groups, pending.uploaded, false);
  }

  async function discardPendingUpload() {
    var pending = state.pendingUpload;
    if (!pending || pending.busy) return;
    var ok = await confirmDialog({
      title: '放弃待上传附件？',
      message: '记录 #' + pending.recordId + ' 已保存，但所选附件将不会上传。此操作不可撤销。',
      confirmText: '放弃附件',
      cancelText: '保留',
      danger: true
    });
    if (!ok) return;
    state.pendingUpload = null;
    renderDetailAlert();
    showToast('已放弃待上传附件，记录 #' + pending.recordId + ' 保留。', 'info');
  }

  async function removeAttachment(recordId, attachmentId, name) {
    if (state.submitting || (state.pendingUpload && state.pendingUpload.busy)) {
      showToast('请等待当前上传或保存操作完成。', 'info');
      return;
    }
    var test = selectedTest();
    if (!test) return;
    var ok = await confirmDialog({
      title: '移除附件？',
      message: '将从记录 #' + recordId + ' 中永久删除“' + name + '”。服务器上的文件也会被删除，此操作不可撤销。',
      confirmText: '移除附件',
      cancelText: '取消',
      danger: true
    });
    if (!ok) return;
    try {
      await apiJson('/api/attachments/' + encodeURIComponent(attachmentId), { method: 'DELETE', timeout: 30000 });
      removeAttachmentLocal(test.id, recordId, attachmentId);
      showToast('附件已移除。', 'success');
    } catch (err) {
      showToast('附件移除失败：' + errorMessage(err), 'error', 6500);
    }
  }

  /* ------------------------------------------------------------------------
     file selection
     ------------------------------------------------------------------------ */

  function addFiles(kind, fileList) {
    var form = state.form;
    if (!form || form.submitting) return;
    var files = Array.prototype.slice.call(fileList || []);
    if (!files.length) return;
    var accepted = 0;
    var rejected = [];
    files.forEach(function (file) {
      var isImage = kind === 'image';
      if (isImage && !fileNameIsImage(file.name)) {
        rejected.push(file.name + '：仅支持 jpg / jpeg 图片');
        return;
      }
      if (!isImage && !fileNameIsData(file.name)) {
        rejected.push(file.name + '：仅支持 xlsx / xls / xlsm');
        return;
      }
      if (!file.size) {
        rejected.push(file.name + '：文件为空');
        return;
      }
      if (file.size > MAX_FILE_BYTES) {
        rejected.push(file.name + '：超过 50 MiB');
        return;
      }
      var list = isImage ? form.imageFiles : form.dataFiles;
      var duplicate = list.some(function (existing) {
        return existing.name === file.name && existing.size === file.size && existing.lastModified === file.lastModified;
      });
      if (duplicate) {
        rejected.push(file.name + '：已在列表中');
        return;
      }
      list.push(file);
      if (isImage) form.previews.set(file, URL.createObjectURL(file));
      accepted += 1;
    });

    if (kind === 'image') {
      var grid = $('#imageFileGrid');
      if (grid) grid.innerHTML = imageFilesHTML(form);
    } else {
      var listHost = $('#dataFileList');
      if (listHost) listHost.innerHTML = dataFilesHTML(form);
    }
    if (rejected.length) {
      showToast(rejected.slice(0, 3).join('；') + (rejected.length > 3 ? ' 等' : ''), accepted ? 'warning' : 'error', 6500);
    }
  }

  function removeFile(kind, index) {
    var form = state.form;
    if (!form || form.submitting) return;
    var list = kind === 'image' ? form.imageFiles : form.dataFiles;
    if (index < 0 || index >= list.length) return;
    var file = list[index];
    if (kind === 'image') {
      var url = form.previews.get(file);
      if (url) URL.revokeObjectURL(url);
      form.previews.delete(file);
    }
    list.splice(index, 1);
    if (kind === 'image') {
      var grid = $('#imageFileGrid');
      if (grid) grid.innerHTML = imageFilesHTML(form);
    } else {
      var host = $('#dataFileList');
      if (host) host.innerHTML = dataFilesHTML(form);
    }
  }

  /* ------------------------------------------------------------------------
     page level actions
     ------------------------------------------------------------------------ */

  async function attemptSelect(id) {
    if (!id) return;
    if (id === state.selectedId) {
      document.body.classList.add('mobile-detail');
      renderTestList();
      return;
    }
    if (state.submitting || (state.pendingUpload && state.pendingUpload.busy)) {
      showToast('正在保存或上传，请稍候再切换。', 'info');
      return;
    }
    var ok = await guardNavigation();
    if (!ok) return;
    state.selectedId = id;
    state.recordFilters = Object.assign({}, RECORD_FILTER_DEFAULT);
    document.body.classList.add('mobile-detail');
    renderTestList();
    renderDetail(true);
    ensureRecordsForSelected(id);
    if (window.matchMedia('(max-width: 900px)').matches) window.scrollTo(0, 0);
  }

  async function backToList() {
    if (state.submitting || (state.pendingUpload && state.pendingUpload.busy)) {
      showToast('正在保存或上传，请稍候。', 'info');
      return;
    }
    var ok = await guardNavigation();
    if (!ok) return;
    renderTestList();
    renderDetail();
    document.body.classList.remove('mobile-detail');
    window.scrollTo(0, 0);
  }

  function setModuleFilter(value) {
    state.filters.module = value;
    renderModuleList();
    renderTestList();
  }

  function setPriorityFilter(value) {
    state.filters.priority = value;
    var select = $('#priorityFilter');
    if (select) select.value = value;
    renderTestList();
  }

  function setSearchQuery(value) {
    state.filters.q = value;
    renderTestList();
  }

  function clearAllFilters() {
    state.filters.module = 'all';
    state.filters.priority = 'all';
    state.filters.q = '';
    var search = $('#searchInput');
    if (search) search.value = '';
    var select = $('#priorityFilter');
    if (select) select.value = 'all';
    renderModuleList();
    renderTestList();
  }

  function clearOneFilter(key) {
    if (key === 'q') {
      state.filters.q = '';
      var search = $('#searchInput');
      if (search) search.value = '';
    } else if (key === 'module') {
      state.filters.module = 'all';
      renderModuleList();
    } else if (key === 'priority') {
      state.filters.priority = 'all';
      var select = $('#priorityFilter');
      if (select) select.value = 'all';
    }
    renderTestList();
  }

  async function exportFile(kind) {
    if (state.exporting) return;
    state.exporting = kind;
    updateExportButtons();
    var isXlsx = kind === 'xlsx';
    showToast(isXlsx ? '正在生成 XLSX 导出…' : '正在打包 ZIP 备份…', 'info', 3000);
    try {
      var project = currentProject();
      var suffix = isXlsx ? ('-' + ((project && project.id) || 'project') + '-records.xlsx') : '-backup.zip';
      await downloadBlob(
        isXlsx ? projectApiPath('/export.xlsx') : '/api/backup.zip',
        'testflow' + suffix
      );
      showToast(isXlsx ? 'XLSX 已导出。' : 'ZIP 备份已下载。', 'success');
    } catch (err) {
      showToast((isXlsx ? 'XLSX 导出失败：' : 'ZIP 备份失败：') + errorMessage(err), 'error', 7000);
    } finally {
      state.exporting = null;
      updateExportButtons();
    }
  }

  function updateExportButtons() {
    var xlsx = $('#exportXlsxBtn');
    var zip = $('#exportZipBtn');
    if (xlsx) {
      xlsx.disabled = !!state.exporting;
      xlsx.setAttribute('aria-busy', state.exporting === 'xlsx' ? 'true' : 'false');
    }
    if (zip) {
      zip.disabled = !!state.exporting;
      zip.setAttribute('aria-busy', state.exporting === 'zip' ? 'true' : 'false');
    }
  }

  /* ------------------------------------------------------------------------
     clipboard: copy / cut / paste test items
     ------------------------------------------------------------------------ */

  async function loadClipboard() {
    try {
      var data = await apiJson('/api/clipboard', { timeout: 15000 });
      if (data && Array.isArray(data.items)) state.clipboard = data;
    } catch (err) {
      state.clipboard = { items: [], count: 0 };
    }
    renderClipboardBar();
  }

  function renderClipboardBar() {
    var bar = $('#clipboardBar');
    var label = $('#clipboardText');
    if (!bar || !label) return;
    var items = (state.clipboard && state.clipboard.items) || [];
    if (!items.length) {
      bar.hidden = true;
      label.textContent = '';
      return;
    }
    var first = items[0];
    var title = first.test ? text(first.test.title, '测试项') : (first.missing ? '（已失效）' : '测试项');
    label.textContent = (first.op === 'cut' ? '已剪切：' : '已复制：') +
      text(first.test_id, '') + ' · ' + title +
      (items.length > 1 ? '（共 ' + items.length + ' 项）' : '');
    bar.hidden = false;
  }

  async function clipboardAction(op) {
    var test = selectedTest();
    if (!test || !state.projectId) return;
    var records = (state.index.get(test.id) || []).length;
    if (op === 'cut' && records) {
      var ok = await confirmDialog({
        title: '剪切测试项',
        danger: true,
        message: '将剪切 ' + test.id + '「' + text(test.title, '') + '」，包含 ' + records + ' 条记录。',
        note: '粘贴到其它项目或模块时，这些记录和附件会一起移动过去。',
        confirmText: '剪切'
      });
      if (!ok) return;
    }
    try {
      var data = await apiJson('/api/clipboard', {
        method: 'POST',
        json: { project_id: state.projectId, test_ids: [test.id], op: op },
        timeout: 15000
      });
      state.clipboard = data && Array.isArray(data.items) ? data : { items: [], count: 0 };
      renderClipboardBar();
      showToast((op === 'copy' ? '已复制 ' : '已剪切 ') + test.id + '，切换项目/模块后点“粘贴”。', 'success', 6000);
    } catch (err) {
      showToast((op === 'copy' ? '复制' : '剪切') + '失败：' + errorMessage(err), 'error', 7000);
    }
  }

  async function clearClipboard() {
    try {
      await apiJson('/api/clipboard', { method: 'DELETE', timeout: 15000 });
      state.clipboard = { items: [], count: 0 };
      renderClipboardBar();
      showToast('已清空剪贴板。', 'info');
    } catch (err) {
      showToast('清空剪贴板失败：' + errorMessage(err), 'error', 7000);
    }
  }

  function openPasteDialog() {
    var items = (state.clipboard && state.clipboard.items) || [];
    if (!items.length) {
      showToast('剪贴板为空，请先复制或剪切测试项。', 'info');
      return Promise.resolve(false);
    }
    var root = $('#modalRoot');
    if (!root) return Promise.resolve(false);
    var previousFocus = document.activeElement;
    var options = state.projects.map(function (project) {
      return '<option value="' + esc(project.id) + '"' + (project.id === state.projectId ? ' selected' : '') + '>' +
        esc(project.name) + '（' + project.test_count + ' 项）</option>';
    }).join('');

    return new Promise(function (resolve) {
      root.innerHTML = '';
      var backdrop = document.createElement('div');
      backdrop.className = 'modal-backdrop';
      var modal = document.createElement('div');
      modal.className = 'modal paste-modal';
      modal.setAttribute('role', 'dialog');
      modal.setAttribute('aria-modal', 'true');
      modal.setAttribute('aria-labelledby', 'pasteModalTitle');
      modal.innerHTML =
        '<div class="modal-head"><span class="modal-badge">' + icon('paste') + '</span>' +
        '<h3 id="pasteModalTitle">粘贴测试项</h3></div>' +
        '<div class="modal-body">' +
          '<p class="modal-note" style="margin-top:0">' +
            esc(items.length > 1 ? ('共 ' + items.length + ' 个测试项，模式：' + (items[0].op === 'cut' ? '剪切' : '复制')) :
              ((items[0].op === 'cut' ? '剪切：' : '复制：') + items[0].test_id + ' · ' + text(items[0].test && items[0].test.title, ''))) +
          '</p>' +
          '<form class="paste-form" novalidate>' +
            '<label class="field"><span class="field-label">目标项目</span>' +
              '<select name="project">' + options + '</select></label>' +
            '<label class="field"><span class="field-label">目标模块</span>' +
              '<input type="text" name="module" maxlength="60" autocomplete="off" placeholder="留空则保持原模块"></label>' +
            '<p class="form-hint">剪切会在粘贴成功后从原位置移除（记录与附件随之移动）；复制则保留原测试项。</p>' +
            '<p class="form-alert" hidden></p>' +
          '</form>' +
        '</div>' +
        '<div class="modal-actions">' +
          '<button type="button" class="btn" data-modal-cancel>取消</button>' +
          '<button type="button" class="btn btn-primary" data-modal-ok>粘贴</button>' +
        '</div>';
      backdrop.appendChild(modal);
      root.appendChild(backdrop);

      var form = modal.querySelector('form');
      var alertBox = modal.querySelector('.form-alert');
      var cancelBtn = modal.querySelector('[data-modal-cancel]');
      var okBtn = modal.querySelector('[data-modal-ok]');
      var closed = false;

      function close(result) {
        if (closed) return;
        closed = true;
        document.removeEventListener('keydown', onKey, true);
        root.innerHTML = '';
        if (previousFocus && typeof previousFocus.focus === 'function') {
          try { previousFocus.focus({ preventScroll: true }); } catch (err) { /* ignore */ }
        }
        resolve(result);
      }

      function onKey(event) {
        if (event.key === 'Escape') {
          event.preventDefault();
          close(false);
        }
      }

      cancelBtn.addEventListener('click', function () { close(false); });
      backdrop.addEventListener('mousedown', function (event) {
        if (event.target === backdrop) close(false);
      });
      document.addEventListener('keydown', onKey, true);
      form.querySelector('input[name="module"]').focus();

      okBtn.addEventListener('click', async function () {
        var payload = {
          project_id: form.querySelector('select[name="project"]').value,
          module: form.querySelector('input[name="module"]').value.trim()
        };
        if (alertBox) alertBox.hidden = true;
        okBtn.disabled = true;
        cancelBtn.disabled = true;
        okBtn.textContent = '粘贴中…';
        try {
          var result = await apiJson('/api/clipboard/paste', {
            method: 'POST', json: payload, timeout: 25000
          });
          if (!result || !Array.isArray(result.pasted)) {
            throw new RequestError('服务器未返回粘贴结果。', 0, 'bad-response');
          }
          state.clipboard = result.clipboard || { items: [], count: 0 };
          renderClipboardBar();
          close(true);
          if (!result.pasted.length) {
            showToast('没有可粘贴的测试项。', 'info');
            return;
          }
          var first = result.pasted[0];
          var targetProject = state.projects.filter(function (p) { return p.id === first.project_id; })[0];
          showToast('已粘贴 ' + result.pasted.length + ' 个测试项到「' +
            (targetProject ? targetProject.name : first.project_id) + '」。', 'success', 6000);
          await focusTest(first.project_id, first.test_id);
        } catch (err) {
          okBtn.disabled = false;
          cancelBtn.disabled = false;
          okBtn.textContent = '粘贴';
          if (alertBox) {
            alertBox.textContent = errorMessage(err);
            alertBox.hidden = false;
          }
        }
      });
    });
  }

  async function focusTest(projectId, testId) {
    if (projectId && projectId !== state.projectId) {
      await loadProjects(projectId);
    }
    await loadBootstrap();
    var exists = state.tests.some(function (test) { return test.id === testId; });
    if (!exists) return;
    state.selectedId = testId;
    document.body.classList.add('mobile-detail');
    renderTestList();
    renderDetail(true);
    ensureRecordsForSelected(testId);
  }

  /* ------------------------------------------------------------------------
     add / delete test items
     ------------------------------------------------------------------------ */

  function nextTestIdSuggestion() {
    var best = null;
    state.tests.forEach(function (test) {
      var match = /^([A-Za-z][A-Za-z0-9_-]*?)(\d+)$/.exec(test.id || '');
      if (!match) return;
      var number = parseInt(match[2], 10);
      if (!best || best.prefix !== match[1]) {
        if (!best) best = { prefix: match[1], max: number, width: match[2].length };
        return;
      }
      if (number > best.max) best.max = number;
      if (match[2].length > best.width) best.width = match[2].length;
    });
    if (!best) return 'T01';
    var next = String(best.max + 1);
    while (next.length < best.width) next = '0' + next;
    return best.prefix + next;
  }

  async function reloadTests(selectId) {
    await loadBootstrap();
    if (!selectId) return;
    var exists = state.tests.some(function (test) { return test.id === selectId; });
    if (!exists) return;
    state.selectedId = selectId;
    document.body.classList.add('mobile-detail');
    renderTestList();
    renderDetail(true);
    ensureRecordsForSelected(selectId);
  }

  function openNewTestDialog() {
    if (!state.projectId) return Promise.resolve(false);
    var root = $('#modalRoot');
    if (!root) return Promise.resolve(false);
    var previousFocus = document.activeElement;
    var guidance = [
      { key: 'purpose', label: '测试目的', rows: 2 },
      { key: 'stimulus', label: '测试连接/激励', rows: 2 },
      { key: 'metrics', label: '需记录的指标', rows: 2 },
      { key: 'conditions', label: '建议测试条件', rows: 2 },
      { key: 'calculation', label: '计算/处理方法', rows: 2 },
      { key: 'criteria', label: '论文参考/建议判据', rows: 2 },
      { key: 'instruments', label: '仪器', rows: 0 }
    ];

    return new Promise(function (resolve) {
      root.innerHTML = '';
      var backdrop = document.createElement('div');
      backdrop.className = 'modal-backdrop';
      var modal = document.createElement('div');
      modal.className = 'modal new-test-modal';
      modal.setAttribute('role', 'dialog');
      modal.setAttribute('aria-modal', 'true');
      modal.setAttribute('aria-labelledby', 'newTestModalTitle');

      var guidanceHtml = guidance.map(function (field) {
        var control = field.rows
          ? '<textarea name="' + field.key + '" rows="' + field.rows + '" maxlength="4000" placeholder="未填写"></textarea>'
          : '<input type="text" name="' + field.key + '" maxlength="4000" autocomplete="off" placeholder="未填写">';
        return '<label class="field"><span class="field-label">' + esc(field.label) + '</span>' + control + '</label>';
      }).join('');

      modal.innerHTML =
        '<div class="modal-head">' +
          '<span class="modal-badge">' + icon('plus') + '</span>' +
          '<h3 id="newTestModalTitle">新建测试项</h3>' +
        '</div>' +
        '<div class="modal-body">' +
          '<form class="new-test-form" novalidate>' +
            '<div class="field-row">' +
              '<label class="field"><span class="field-label">测试编号</span>' +
                '<input type="text" name="id" maxlength="32" autocomplete="off" placeholder="留空自动生成" value="' + esc(nextTestIdSuggestion()) + '"></label>' +
              '<label class="field"><span class="field-label">优先级</span>' +
                '<input type="text" name="priority" maxlength="20" autocomplete="off" placeholder="P0 / P1 / P2" value="P0"></label>' +
            '</div>' +
            '<label class="field"><span class="field-label">模块</span>' +
              '<input type="text" name="module" maxlength="60" autocomplete="off" placeholder="例如：TIA（留空为“未分组”）"></label>' +
            '<label class="field"><span class="field-label">测试项目<span class="req">*</span></span>' +
              '<input type="text" name="title" maxlength="200" autocomplete="off" placeholder="例如：跨阻增益"></label>' +
            '<details class="new-test-guidance"><summary>测试指导（可选）</summary>' + guidanceHtml + '</details>' +
            '<p class="form-hint">新建的测试项会追加到当前项目末尾；测试指导之后也能用“编辑说明”修改。</p>' +
            '<p class="form-alert" hidden></p>' +
          '</form>' +
        '</div>' +
        '<div class="modal-actions">' +
          '<button type="button" class="btn" data-modal-cancel>取消</button>' +
          '<button type="button" class="btn btn-primary" data-modal-ok>创建测试项</button>' +
        '</div>';
      backdrop.appendChild(modal);
      root.appendChild(backdrop);

      var form = modal.querySelector('form');
      var alertBox = modal.querySelector('.form-alert');
      var cancelBtn = modal.querySelector('[data-modal-cancel]');
      var okBtn = modal.querySelector('[data-modal-ok]');
      var titleInput = form.querySelector('[name="title"]');
      var closed = false;

      function showAlert(message) {
        if (!alertBox) return;
        alertBox.textContent = message;
        alertBox.hidden = false;
      }

      function close(result) {
        if (closed) return;
        closed = true;
        document.removeEventListener('keydown', onKey, true);
        root.innerHTML = '';
        if (previousFocus && typeof previousFocus.focus === 'function') {
          try { previousFocus.focus({ preventScroll: true }); } catch (err) { /* ignore */ }
        }
        resolve(result);
      }

      function onKey(event) {
        if (event.key === 'Escape') {
          event.preventDefault();
          close(false);
        }
      }

      cancelBtn.addEventListener('click', function () { close(false); });
      backdrop.addEventListener('mousedown', function (event) {
        if (event.target === backdrop) close(false);
      });
      document.addEventListener('keydown', onKey, true);
      titleInput.focus();

      okBtn.addEventListener('click', async function () {
        var payload = {};
        ['id', 'priority', 'module', 'title'].concat(guidance.map(function (f) { return f.key; }))
          .forEach(function (key) {
            var input = form.querySelector('[name="' + key + '"]');
            payload[key] = input ? input.value : '';
          });
        if (!payload.title.trim()) {
          showAlert('请填写测试项目名称。');
          titleInput.focus();
          return;
        }
        if (alertBox) alertBox.hidden = true;
        okBtn.disabled = true;
        cancelBtn.disabled = true;
        okBtn.textContent = '创建中…';
        try {
          var result = await apiJson(projectApiPath('/tests'), {
            method: 'POST', json: payload, timeout: 20000
          });
          if (!result || !result.test) throw new RequestError('服务器未返回新建的测试项。', 0, 'bad-response');
          close(true);
          showToast('已创建测试项 ' + result.test.id + '「' + text(result.test.title, '') + '」。', 'success', 6000);
          await reloadTests(result.test.id);
        } catch (err) {
          okBtn.disabled = false;
          cancelBtn.disabled = false;
          okBtn.textContent = '创建测试项';
          showAlert(errorMessage(err));
        }
      });
    });
  }

  async function deleteCurrentTest() {
    var test = selectedTest();
    if (!test || !state.projectId) return;
    if (state.form && state.form.testId === test.id && isFormDirty()) {
      showToast('请先保存或取消正在编辑的记录。', 'info');
      return;
    }
    var records = (state.index.get(test.id) || []).length;
    var ok = await confirmDialog({
      title: '删除测试项',
      danger: true,
      message: '确定删除 ' + test.id + '「' + text(test.title, '未命名测试项') + '」吗？',
      note: records
        ? '该测试项下的 ' + records + ' 条记录及其附件会一并删除，且无法恢复。'
        : '该测试项下暂无记录；删除后无法恢复。所属项目与其它测试项不受影响。',
      confirmText: '删除测试项'
    });
    if (!ok) return;
    try {
      await apiJson(projectApiPath('/tests/' + encodeURIComponent(test.id)), {
        method: 'DELETE', timeout: 20000
      });
      showToast('已删除测试项 ' + test.id + '。', 'success');
      await reloadTests(null);
    } catch (err) {
      showToast('删除测试项失败：' + errorMessage(err), 'error', 7000);
    }
  }

  /* ------------------------------------------------------------------------
     editable test guidance
     ------------------------------------------------------------------------ */

  function openGuidanceDialog() {
    var test = selectedTest();
    if (!test || !state.projectId) return Promise.resolve(false);
    var root = $('#modalRoot');
    if (!root) return Promise.resolve(false);
    var previousFocus = document.activeElement;
    var formFields = [
      { key: 'purpose', label: '测试目的', rows: 2 },
      { key: 'stimulus', label: '测试连接/激励', rows: 3 },
      { key: 'metrics', label: '需记录的指标', rows: 3 },
      { key: 'conditions', label: '建议测试条件', rows: 3 },
      { key: 'calculation', label: '计算/处理方法', rows: 3 },
      { key: 'criteria', label: '论文参考/建议判据', rows: 3 },
      { key: 'instruments', label: '仪器', rows: 0 }
    ];

    return new Promise(function (resolve) {
      root.innerHTML = '';
      var backdrop = document.createElement('div');
      backdrop.className = 'modal-backdrop';
      var modal = document.createElement('div');
      modal.className = 'modal guidance-modal';
      modal.setAttribute('role', 'dialog');
      modal.setAttribute('aria-modal', 'true');
      modal.setAttribute('aria-labelledby', 'guidanceModalTitle');

      var body = formFields.map(function (field) {
        var value = text(test[field.key]);
        var control = field.rows
          ? '<textarea name="' + field.key + '" rows="' + field.rows + '" maxlength="4000" placeholder="未填写">' + esc(value) + '</textarea>'
          : '<input type="text" name="' + field.key + '" maxlength="4000" autocomplete="off" placeholder="未填写" value="' + esc(value) + '">';
        return '<label class="field"><span class="field-label">' + esc(field.label) + '</span>' + control + '</label>';
      }).join('');

      modal.innerHTML =
        '<div class="modal-head">' +
          '<span class="modal-badge">' + icon('edit') + '</span>' +
          '<h3 id="guidanceModalTitle">编辑测试说明</h3>' +
        '</div>' +
        '<div class="modal-body">' +
          '<p class="modal-note" style="margin-top:0">' + esc(test.id) + ' · ' + esc(text(test.title, '')) + '</p>' +
          '<form class="guidance-form" novalidate>' + body +
            '<p class="form-hint">这些内容对应导入表格中的同名列；修改只影响当前项目，不会改写原始 Excel。</p>' +
            '<p class="form-alert" hidden></p>' +
          '</form>' +
        '</div>' +
        '<div class="modal-actions">' +
          '<button type="button" class="btn" data-modal-cancel>取消</button>' +
          '<button type="button" class="btn btn-primary" data-modal-ok>保存</button>' +
        '</div>';
      backdrop.appendChild(modal);
      root.appendChild(backdrop);

      var form = modal.querySelector('form');
      var alertBox = modal.querySelector('.form-alert');
      var cancelBtn = modal.querySelector('[data-modal-cancel]');
      var okBtn = modal.querySelector('[data-modal-ok]');
      var closed = false;

      function close(result) {
        if (closed) return;
        closed = true;
        document.removeEventListener('keydown', onKey, true);
        root.innerHTML = '';
        if (previousFocus && typeof previousFocus.focus === 'function') {
          try { previousFocus.focus({ preventScroll: true }); } catch (err) { /* ignore */ }
        }
        resolve(result);
      }

      function onKey(event) {
        if (event.key === 'Escape') {
          event.preventDefault();
          close(false);
        }
      }

      cancelBtn.addEventListener('click', function () { close(false); });
      backdrop.addEventListener('mousedown', function (event) {
        if (event.target === backdrop) close(false);
      });
      document.addEventListener('keydown', onKey, true);
      var first = form.querySelector('textarea, input');
      if (first) first.focus();

      okBtn.addEventListener('click', async function () {
        var payload = {};
        formFields.forEach(function (field) {
          var input = form.querySelector('[name="' + field.key + '"]');
          payload[field.key] = input ? input.value : '';
        });
        if (alertBox) alertBox.hidden = true;
        okBtn.disabled = true;
        cancelBtn.disabled = true;
        okBtn.textContent = '保存中…';
        try {
          var result = await apiJson(
            projectApiPath('/tests/' + encodeURIComponent(test.id)),
            { method: 'PUT', json: payload, timeout: 20000 }
          );
          if (!result || !result.test) throw new RequestError('服务器未返回更新后的测试项。', 0, 'bad-response');
          var target = state.tests.filter(function (t) { return t.id === test.id; })[0];
          if (target) {
            Object.keys(result.test).forEach(function (key) { target[key] = result.test[key]; });
          }
          close(true);
          renderTestList();
          renderDetail();
          showToast('已保存 ' + test.id + ' 的测试说明。', 'success');
        } catch (err) {
          okBtn.disabled = false;
          cancelBtn.disabled = false;
          okBtn.textContent = '保存';
          if (alertBox) {
            alertBox.textContent = errorMessage(err);
            alertBox.hidden = false;
          }
        }
      });
    });
  }

  /* ------------------------------------------------------------------------
     project switching / creation
     ------------------------------------------------------------------------ */

  async function loadProjects(preferredId) {
    var data = await apiJson('/api/projects', { timeout: 25000 });
    if (!data || !Array.isArray(data.projects)) {
      throw new RequestError('服务器返回的项目列表格式异常。', 0, 'bad-response');
    }
    state.projects = data.projects.filter(function (p) {
      return p && typeof p.id === 'string' && p.id !== '';
    });
    if (!state.projects.length) {
      state.projectId = null;
      state.project = null;
      syncProjectUrl();
      renderProjectSelect();
      return null;
    }
    var wanted = preferredId || state.projectId;
    var match = state.projects.filter(function (p) { return p.id === wanted; })[0];
    state.projectId = match ? match.id : state.projects[0].id;
    state.project = state.projects.filter(function (p) { return p.id === state.projectId; })[0] || null;
    syncProjectUrl();
    renderProjectSelect();
    return state.projectId;
  }

  function renderProjectSelect() {
    var select = $('#projectSelect');
    var del = $('#deleteProjectBtn');
    if (del) del.disabled = state.projects.length <= 1;
    if (!select) return;
    if (!state.projects.length) {
      select.innerHTML = '<option value="">（还没有项目）</option>';
      select.disabled = true;
      return;
    }
    select.disabled = false;
    select.innerHTML = state.projects.map(function (p) {
      var label = p.name + '（' + p.test_count + ' 项 / ' + p.record_count + ' 条记录）';
      return '<option value="' + esc(p.id) + '"' + (p.id === state.projectId ? ' selected' : '') + '>' +
        esc(label) + '</option>';
    }).join('');
    select.value = state.projectId || '';
  }

  async function switchProject(projectId) {
    if (!projectId || projectId === state.projectId) return;
    if (state.submitting || (state.pendingUpload && state.pendingUpload.busy)) {
      showToast('正在保存或上传，请稍候再切换项目。', 'info');
      renderProjectSelect();
      return;
    }
    var ok = await guardNavigation();
    if (!ok) {
      renderProjectSelect();
      return;
    }
    state.projectId = projectId;
    state.project = state.projects.filter(function (p) { return p.id === projectId; })[0] || null;
    syncProjectUrl();
    renderProjectSelect();
    await loadBootstrap();
  }

  async function deleteCurrentProject() {
    var project = currentProject();
    if (!project) return;
    if (state.projects.length <= 1) {
      showToast('至少需要保留一个项目。', 'info');
      return;
    }
    if (state.submitting || (state.pendingUpload && state.pendingUpload.busy)) {
      showToast('正在保存或上传，请稍候。', 'info');
      return;
    }
    var ok = await confirmDialog({
      title: '删除项目',
      danger: true,
      message: '确定要删除项目「' + project.name + '」吗？',
      note: '该项目下的全部测试记录和附件文件都会被删除，且无法恢复。',
      confirmText: '删除项目'
    });
    if (!ok) return;
    try {
      await apiJson(projectApiPath(''), { method: 'DELETE', timeout: 30000 });
      showToast('已删除项目「' + project.name + '」。', 'success');
      state.projectId = null;
      state.project = null;
      await loadProjects();
      await loadBootstrap();
    } catch (err) {
      showToast('删除项目失败：' + errorMessage(err), 'error', 7000);
    }
  }

  function openNewProjectDialog() {
    var root = $('#modalRoot');
    if (!root) return Promise.resolve(false);
    var previousFocus = document.activeElement;
    return new Promise(function (resolve) {
      root.innerHTML = '';
      var backdrop = document.createElement('div');
      backdrop.className = 'modal-backdrop';
      var modal = document.createElement('div');
      modal.className = 'modal project-modal';
      modal.setAttribute('role', 'dialog');
      modal.setAttribute('aria-modal', 'true');
      modal.setAttribute('aria-labelledby', 'projectModalTitle');
      modal.innerHTML =
        '<div class="modal-head">' +
          '<span class="modal-badge">' + icon('upload') + '</span>' +
          '<h3 id="projectModalTitle">新建测试项目</h3>' +
        '</div>' +
        '<div class="modal-body">' +
          '<form class="project-form" novalidate>' +
            '<label class="field">' +
              '<span class="field-label">项目名称<span class="label-hint">留空则使用文件名</span></span>' +
              '<input type="text" name="name" maxlength="80" autocomplete="off" placeholder="例如：AFE 推荐测试项目">' +
            '</label>' +
            '<label class="field">' +
              '<span class="field-label">测试项目表格（.xlsx / .xlsm）</span>' +
              '<input type="file" name="workbook" accept=".xlsx,.xlsm">' +
            '</label>' +
            '<p class="form-hint">表格结构与《AFE_推荐测试项目与结果记录表1.xlsx》一致：测试计划表需要「序号 / 优先级 / 模块 / 测试项目 / 测试目的」等表头；“指标汇总”“填写说明”为可选工作表，列顺序可以不同。</p>' +
            '<p class="form-alert" hidden></p>' +
          '</form>' +
        '</div>' +
        '<div class="modal-actions">' +
          '<button type="button" class="btn" data-modal-cancel>取消</button>' +
          '<button type="button" class="btn btn-primary" data-modal-ok>创建项目</button>' +
        '</div>';
      backdrop.appendChild(modal);
      root.appendChild(backdrop);

      var form = modal.querySelector('form');
      var alertBox = modal.querySelector('.form-alert');
      var nameInput = form.querySelector('input[name="name"]');
      var fileInput = form.querySelector('input[name="workbook"]');
      var cancelBtn = modal.querySelector('[data-modal-cancel]');
      var okBtn = modal.querySelector('[data-modal-ok]');
      var closed = false;

      function showAlert(message) {
        if (!alertBox) return;
        alertBox.textContent = message;
        alertBox.hidden = false;
      }

      function close(result) {
        if (closed) return;
        closed = true;
        document.removeEventListener('keydown', onKey, true);
        root.innerHTML = '';
        if (previousFocus && typeof previousFocus.focus === 'function') {
          try { previousFocus.focus({ preventScroll: true }); } catch (err) { /* ignore */ }
        }
        resolve(result);
      }

      function onKey(event) {
        if (event.key === 'Escape') {
          event.preventDefault();
          close(false);
        }
      }

      cancelBtn.addEventListener('click', function () { close(false); });
      backdrop.addEventListener('mousedown', function (event) {
        if (event.target === backdrop) close(false);
      });
      document.addEventListener('keydown', onKey, true);
      nameInput.focus();

      okBtn.addEventListener('click', async function () {
        var file = fileInput.files && fileInput.files[0];
        if (!file) {
          showAlert('请选择 .xlsx 或 .xlsm 格式的测试项目表格。');
          fileInput.focus();
          return;
        }
        if (!/\.(xlsx|xlsm)$/i.test(file.name)) {
          showAlert('仅支持 .xlsx / .xlsm；.xls 请在 Excel 中另存为 .xlsx 后再上传。');
          fileInput.focus();
          return;
        }
        if (alertBox) alertBox.hidden = true;
        okBtn.disabled = true;
        cancelBtn.disabled = true;
        okBtn.textContent = '正在解析…';
        var payload = new FormData();
        payload.append('name', nameInput.value.trim());
        payload.append('file', file, file.name);
        try {
          var res = await fetch('/api/projects', {
            method: 'POST',
            body: payload,
            credentials: 'same-origin',
            cache: 'no-store'
          });
          var data = null;
          try { data = await res.json(); } catch (err) { data = null; }
          if (!res.ok) {
            throw new Error((data && data.error) || ('创建项目失败（HTTP ' + res.status + '）'));
          }
          close(true);
          showToast('已创建项目「' + data.project.name + '」，共 ' + data.tests + ' 个测试项。', 'success', 6000);
          await loadProjects(data.project.id);
          await loadBootstrap();
        } catch (err) {
          okBtn.disabled = false;
          cancelBtn.disabled = false;
          okBtn.textContent = '创建项目';
          showAlert(errorMessage(err));
        }
      });
    });
  }

  /* ------------------------------------------------------------------------
     toast / modal / lightbox
     ------------------------------------------------------------------------ */

  function showToast(message, type, duration) {
    var host = $('#toastArea');
    if (!host) return;
    type = type || 'info';
    var icons = { success: 'check', error: 'alert', warning: 'alert', info: 'info' };
    var toast = document.createElement('div');
    toast.className = 'toast toast-' + type;
    toast.innerHTML = '<span class="toast-icon">' + icon(icons[type] || 'info') + '</span>' +
      '<span class="toast-text">' + esc(message) + '</span>';
    host.appendChild(toast);
    while (host.children.length > 4) {
      host.removeChild(host.firstElementChild);
    }
    var life = duration || (type === 'error' ? 6500 : 4200);
    setTimeout(function () {
      if (!toast.parentNode) return;
      toast.classList.add('leaving');
      setTimeout(function () {
        if (toast.parentNode) toast.parentNode.removeChild(toast);
      }, 180);
    }, life);
  }

  function confirmDialog(options) {
    options = options || {};
    if (activeModalClose) {
      var previous = activeModalClose;
      activeModalClose = null;
      previous(false);
    }
    return new Promise(function (resolve) {
      var root = $('#modalRoot');
      var previousFocus = document.activeElement;
      root.innerHTML = '';
      var backdrop = document.createElement('div');
      backdrop.className = 'modal-backdrop';

      var modal = document.createElement('div');
      modal.className = 'modal';
      modal.setAttribute('role', 'dialog');
      modal.setAttribute('aria-modal', 'true');
      modal.setAttribute('aria-labelledby', 'modalTitle');

      var head = document.createElement('div');
      head.className = 'modal-head';
      var badge = document.createElement('span');
      badge.className = 'modal-badge' + (options.danger ? ' danger' : '');
      badge.innerHTML = icon(options.danger ? 'alert' : 'info');
      var title = document.createElement('h3');
      title.id = 'modalTitle';
      title.textContent = options.title || '请确认';
      head.appendChild(badge);
      head.appendChild(title);

      var body = document.createElement('div');
      body.className = 'modal-body';
      var p = document.createElement('p');
      p.textContent = options.message || '';
      body.appendChild(p);
      if (options.note) {
        var note = document.createElement('p');
        note.className = 'modal-note';
        note.textContent = options.note;
        body.appendChild(note);
      }

      var actions = document.createElement('div');
      actions.className = 'modal-actions';
      var cancel = document.createElement('button');
      cancel.type = 'button';
      cancel.className = 'btn';
      cancel.textContent = options.cancelText || '取消';
      var confirm = document.createElement('button');
      confirm.type = 'button';
      confirm.className = 'btn ' + (options.danger ? 'btn-danger' : 'btn-primary');
      confirm.textContent = options.confirmText || '确认';

      actions.appendChild(cancel);
      actions.appendChild(confirm);
      modal.appendChild(head);
      modal.appendChild(body);
      modal.appendChild(actions);
      backdrop.appendChild(modal);
      root.appendChild(backdrop);

      var closed = false;
      function close(value) {
        if (closed) return;
        closed = true;
        document.removeEventListener('keydown', onKey, true);
        if (activeModalClose === close) activeModalClose = null;
        root.innerHTML = '';
        if (previousFocus && typeof previousFocus.focus === 'function') {
          try { previousFocus.focus({ preventScroll: true }); } catch (err) { /* ignore */ }
        }
        resolve(value);
      }
      activeModalClose = close;

      function onKey(event) {
        if (event.key === 'Escape') {
          event.preventDefault();
          close(false);
        } else if (event.key === 'Tab') {
          var focusable = [cancel, confirm];
          var index = focusable.indexOf(document.activeElement);
          if (index < 0) {
            event.preventDefault();
            cancel.focus();
          } else if (event.shiftKey && index === 0) {
            event.preventDefault();
            confirm.focus();
          } else if (!event.shiftKey && index === focusable.length - 1) {
            event.preventDefault();
            cancel.focus();
          }
        }
      }

      cancel.addEventListener('click', function () { close(false); });
      confirm.addEventListener('click', function () { close(true); });
      backdrop.addEventListener('mousedown', function (event) {
        if (event.target === backdrop) close(false);
      });
      document.addEventListener('keydown', onKey, true);
      cancel.focus();
    });
  }

  function openLightbox(url, name) {
    var resolved = safeUrl(url);
    if (!resolved) {
      showToast('图片地址无效，无法预览。', 'error');
      return;
    }
    lightboxReturnFocus = document.activeElement;
    var box = $('#lightbox');
    var img = $('#lightboxImage');
    var caption = $('#lightboxCaption');
    if (!box || !img) return;
    img.src = resolved;
    img.alt = name || '附件预览';
    if (caption) caption.textContent = name || '附件预览';
    box.hidden = false;
    document.body.classList.add('lightbox-open');
    var closeButton = $('#lightboxClose');
    if (closeButton) closeButton.focus();
  }

  function closeLightbox() {
    var box = $('#lightbox');
    if (!box || box.hidden) return;
    box.hidden = true;
    document.body.classList.remove('lightbox-open');
    var img = $('#lightboxImage');
    if (img) img.src = '';
    if (lightboxReturnFocus && typeof lightboxReturnFocus.focus === 'function') {
      try { lightboxReturnFocus.focus({ preventScroll: true }); } catch (err) { /* ignore */ }
    }
    lightboxReturnFocus = null;
  }

  /* ------------------------------------------------------------------------
     event wiring
     ------------------------------------------------------------------------ */

  function bindStaticEvents() {
    var search = $('#searchInput');
    if (search) {
      search.addEventListener('input', function () { setSearchQuery(search.value); });
      search.addEventListener('keydown', function (event) {
        if (event.key === 'Escape' && search.value) {
          search.value = '';
          setSearchQuery('');
        }
      });
    }

    var priority = $('#priorityFilter');
    if (priority) priority.addEventListener('change', function () { setPriorityFilter(priority.value); });

    var clear = $('#clearFiltersBtn');
    if (clear) clear.addEventListener('click', clearAllFilters);

    var modules = $('#moduleList');
    if (modules) {
      modules.addEventListener('click', function (event) {
        var button = event.target.closest('[data-module]');
        if (!button) return;
        setModuleFilter(button.getAttribute('data-module'));
      });
    }

    var list = $('#testList');
    if (list) {
      list.addEventListener('click', function (event) {
        var card = event.target.closest('[data-test-id]');
        if (card) attemptSelect(card.getAttribute('data-test-id'));
      });
    }

    var activeFilters = $('#activeFilters');
    if (activeFilters) {
      activeFilters.addEventListener('click', function (event) {
        var button = event.target.closest('[data-clear-filter]');
        if (button) clearOneFilter(button.getAttribute('data-clear-filter'));
      });
    }

    var detail = $('#detailPane');
    if (detail) {
      detail.addEventListener('click', onDetailClick);
      detail.addEventListener('submit', onDetailSubmit);
      detail.addEventListener('input', onDetailInput);
      detail.addEventListener('change', onDetailChange);
      detail.addEventListener('keydown', onDetailKeydown);
      detail.addEventListener('dragover', onDetailDragOver);
      detail.addEventListener('dragleave', onDetailDragLeave);
      detail.addEventListener('drop', onDetailDrop);
    }

    var projectSelect = $('#projectSelect');
    if (projectSelect) {
      projectSelect.addEventListener('change', function () { switchProject(projectSelect.value); });
    }
    var newProjectBtn = $('#newProjectBtn');
    if (newProjectBtn) newProjectBtn.addEventListener('click', openNewProjectDialog);
    var newTestBtn = $('#newTestBtn');
    if (newTestBtn) newTestBtn.addEventListener('click', openNewTestDialog);
    var pasteBtn = $('#pasteClipboardBtn');
    if (pasteBtn) pasteBtn.addEventListener('click', openPasteDialog);
    var clearClipBtn = $('#clearClipboardBtn');
    if (clearClipBtn) clearClipBtn.addEventListener('click', clearClipboard);
    var deleteProjectBtn = $('#deleteProjectBtn');
    if (deleteProjectBtn) deleteProjectBtn.addEventListener('click', deleteCurrentProject);

    var xlsx = $('#exportXlsxBtn');
    if (xlsx) xlsx.addEventListener('click', function () { exportFile('xlsx'); });
    var zip = $('#exportZipBtn');
    if (zip) zip.addEventListener('click', function () { exportFile('zip'); });

    var retryIndex = $('#retryIndexBtn');
    if (retryIndex) {
      retryIndex.addEventListener('click', function () {
        if (state.indexState === 'loading') return;
        loadRecordIndex({ onlyFailed: state.indexFailed.size > 0 });
      });
    }

    var lightboxClose = $('#lightboxClose');
    if (lightboxClose) lightboxClose.addEventListener('click', closeLightbox);
    var backdrop = document.querySelector('[data-lightbox-close]');
    if (backdrop) backdrop.addEventListener('click', closeLightbox);

    document.addEventListener('click', function (event) {
      var retry = event.target.closest('[data-action="retry-bootstrap"]');
      if (retry && !state.bootstrapping) loadBootstrap();
    });

    document.addEventListener('keydown', function (event) {
      if (event.key === 'Escape') {
        var box = $('#lightbox');
        if (box && !box.hidden) {
          event.preventDefault();
          closeLightbox();
          return;
        }
      }
      if (event.key === '/' && !isTypingTarget(event.target) && !event.metaKey && !event.ctrlKey && !event.altKey) {
        event.preventDefault();
        if (search) search.focus();
      }
    });

    window.addEventListener('beforeunload', function (event) {
      if (isFormDirty() || state.pendingUpload || state.submitting) {
        event.preventDefault();
        event.returnValue = '';
        return '';
      }
      return undefined;
    });
  }

  function onDetailClick(event) {
    var dropzone = event.target.closest('[data-dropzone]');
    if (dropzone && event.target.tagName && event.target.tagName.toLowerCase() !== 'input') {
      if (!state.form || state.form.submitting) return;
      var picker = dropzone.querySelector('input[type="file"]');
      if (picker && !picker.disabled) picker.click();
      return;
    }

    var statusButton = event.target.closest('[data-record-status]');
    if (statusButton) {
      state.recordFilters.status = statusButton.getAttribute('data-record-status');
      renderRecordsArea();
      return;
    }
    var modeButton = event.target.closest('[data-record-mode]');
    if (modeButton) {
      state.recordFilters.mode = modeButton.getAttribute('data-record-mode');
      renderRecordsArea();
      return;
    }

    var target = event.target.closest('[data-action]');
    if (!target) return;
    var action = target.getAttribute('data-action');

    switch (action) {
      case 'new-record':
        openNewForm();
        break;
      case 'edit-guidance':
        openGuidanceDialog();
        break;
      case 'delete-test':
        deleteCurrentTest();
        break;
      case 'copy-test':
        clipboardAction('copy');
        break;
      case 'cut-test':
        clipboardAction('cut');
        break;
      case 'edit-record':
        openEditForm(Number(target.getAttribute('data-record-id')));
        break;
      case 'remove-attachment':
        removeAttachment(
          Number(target.getAttribute('data-record-id')),
          Number(target.getAttribute('data-attachment-id')),
          target.getAttribute('data-attachment-name') || '该附件'
        );
        break;
      case 'cancel-form':
        cancelForm();
        break;
      case 'remove-data-file':
        removeFile('data', Number(target.getAttribute('data-index')));
        break;
      case 'remove-image-file':
        removeFile('image', Number(target.getAttribute('data-index')));
        break;
      case 'retry-records':
      case 'refresh-records':
        refreshSelectedRecords(action === 'retry-records');
        break;
      case 'refresh-from-form':
        refreshFromForm();
        break;
      case 'retry-upload':
        retryPendingUpload();
        break;
      case 'discard-upload':
        discardPendingUpload();
        break;
      case 'clear-record-filters':
        state.recordFilters = Object.assign({}, RECORD_FILTER_DEFAULT);
        renderRecordsArea();
        break;
      case 'back-to-list':
        backToList();
        break;
      case 'open-image':
        openLightbox(target.getAttribute('data-url'), target.getAttribute('data-name') || '附件预览');
        break;
      case 'retry-bootstrap':
        if (!state.bootstrapping) loadBootstrap();
        break;
      case 'clear-all-filters':
        clearAllFilters();
        break;
      default:
        break;
    }
  }

  function onDetailSubmit(event) {
    if (event.target && event.target.id === 'recordForm') {
      event.preventDefault();
      submitForm();
    }
  }

  function onDetailInput(event) {
    var input = event.target.closest('[data-field]');
    if (!input) return;
    setFormValue(input);
  }

  function onDetailChange(event) {
    var fileInput = event.target.closest('[data-file-kind]');
    if (fileInput) {
      var kind = fileInput.getAttribute('data-file-kind');
      addFiles(kind, fileInput.files);
      fileInput.value = '';
      return;
    }
    var input = event.target.closest('[data-field]');
    if (input) setFormValue(input);
  }

  function onDetailKeydown(event) {
    var dropzone = event.target.closest('[data-dropzone]');
    if (dropzone && (event.key === 'Enter' || event.key === ' ' || event.key === 'Spacebar')) {
      event.preventDefault();
      var input = dropzone.querySelector('input[type="file"]');
      if (input) input.click();
    }
  }

  function onDetailDragOver(event) {
    var dropzone = event.target.closest('[data-dropzone]');
    if (!dropzone) return;
    event.preventDefault();
    dropzone.classList.add('dragover');
  }

  function onDetailDragLeave(event) {
    var dropzone = event.target.closest('[data-dropzone]');
    if (!dropzone) return;
    if (dropzone.contains(event.relatedTarget)) return;
    dropzone.classList.remove('dragover');
  }

  function onDetailDrop(event) {
    var dropzone = event.target.closest('[data-dropzone]');
    if (!dropzone) return;
    event.preventDefault();
    dropzone.classList.remove('dragover');
    if (!state.form || state.form.submitting) return;
    addFiles(dropzone.getAttribute('data-dropzone'), event.dataTransfer.files);
  }

  function refreshSelectedRecords(notify) {
    var id = state.selectedId;
    if (!id) return;
    if (state.recordsLoading.has(id)) return;
    refreshRecords(id, { force: true, notify: !!notify }).catch(function (err) {
      if (notify) showToast('记录刷新失败：' + errorMessage(err), 'error', 6500);
    });
  }

  async function refreshFromForm() {
    var form = state.form;
    if (!form) return;
    if (state.recordsLoading.has(form.testId)) return;
    try {
      await refreshRecords(form.testId, { force: true, notify: false });
      if (state.form === form) {
        form.blockedByAmbiguity = false;
        form.alert = {
          type: 'warning',
          title: '记录已刷新，请确认',
          text: '记录历史已重新读取。若列表中已出现刚才保存的记录，请取消表单；若没有，可再次保存。检查后再提交可避免重复记录。'
        };
        renderForm();
      }
      showToast('记录列表已刷新，请确认是否已写入。', 'info', 6000);
    } catch (err) {
      showToast('刷新记录失败：' + errorMessage(err), 'error', 6500);
    }
  }

  function renderAll() {
    renderProjectSelect();
    renderTopStats();
    renderModuleList();
    renderPriorityStats();
    renderRecordStats();
    renderPriorityFilter();
    renderInstructions();
    renderSource();
    renderTestList();
  }

  function renderBootstrapError() {
    renderProjectSelect();
    renderTopStats();
    renderModuleList();
    renderPriorityStats();
    renderRecordStats();
    renderPriorityFilter();
    renderInstructions();
    renderSource();
    renderTestList();
    renderDetail();
  }

  function renderInitial() {
    renderProjectSelect();
    renderTopStats();
    renderModuleList();
    renderPriorityStats();
    renderRecordStats();
    renderPriorityFilter();
    renderInstructions();
    renderSource();
    renderTestList();
    renderDetail();
  }

  init();

  function init() {
    bindStaticEvents();
    renderInitial();
    boot();
  }

  async function boot() {
    try {
      await loadProjects(projectFromUrl());
    } catch (err) {
      state.phase = 'error';
      state.loadError = '无法加载项目列表：' + errorMessage(err);
      renderBootstrapError();
      return;
    }
    if (!state.projectId) {
      state.phase = 'ready';
      state.tests = [];
      state.summary = [];
      state.instructions = [];
      state.source = '';
      renderAll();
      renderDetail();
      await loadClipboard();
      return;
    }
    await loadBootstrap();
    await loadClipboard();
  }
})();
