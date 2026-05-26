/* 在线签名页逻辑；由 sign.html 以 <script src> 引入，勿在聊天里复制粘贴覆盖 */
(function () {
  'use strict';
  // 供 sign_file.html 的 watchSignBoot 区分「脚本未加载」与「异步初始化仍在进行」
  try {
    window.__SIGN_PAGE_JS_EXECUTED = true;
  } catch (_) {}

  // 兼容部分老旧移动端 WebView：避免 Object.assign / Array.from 缺失导致整页脚本中断。
  if (!Object.assign) {
    Object.assign = function (target) {
      if (target == null) throw new TypeError('Cannot convert undefined or null to object');
      var to = Object(target);
      for (var i = 1; i < arguments.length; i++) {
        var src = arguments[i];
        if (src == null) continue;
        for (var key in src) {
          if (Object.prototype.hasOwnProperty.call(src, key)) to[key] = src[key];
        }
      }
      return to;
    };
  }
  if (!Array.from) {
    Array.from = function (arrLike) {
      return Array.prototype.slice.call(arrLike);
    };
  }
  if (typeof window !== 'undefined' && !window.requestAnimationFrame) {
    window.requestAnimationFrame = function (cb) {
      return setTimeout(function () {
        cb(Date.now());
      }, 16);
    };
  }
  if (typeof window !== 'undefined' && !window.cancelAnimationFrame) {
    window.cancelAnimationFrame = function (id) {
      clearTimeout(id);
    };
  }
  if (typeof Promise === 'undefined') {
    (function () {
      function MiniPromise(executor) {
        this._state = 'pending';
        this._value = undefined;
        this._handlers = [];
        var self = this;
        function resolve(v) {
          if (self._state !== 'pending') return;
          self._state = 'fulfilled';
          self._value = v;
          self._flush();
        }
        function reject(e) {
          if (self._state !== 'pending') return;
          self._state = 'rejected';
          self._value = e;
          self._flush();
        }
        try {
          executor(resolve, reject);
        } catch (e) {
          reject(e);
        }
      }
      MiniPromise.prototype._flush = function () {
        var self = this;
        setTimeout(function () {
          while (self._handlers.length) {
            var h = self._handlers.shift();
            try {
              if (self._state === 'fulfilled') {
                if (typeof h.onFulfilled === 'function') {
                  h.resolve(h.onFulfilled(self._value));
                } else {
                  h.resolve(self._value);
                }
              } else if (self._state === 'rejected') {
                if (typeof h.onRejected === 'function') {
                  h.resolve(h.onRejected(self._value));
                } else {
                  h.reject(self._value);
                }
              }
            } catch (e) {
              h.reject(e);
            }
          }
        }, 0);
      };
      MiniPromise.prototype.then = function (onFulfilled, onRejected) {
        var self = this;
        return new MiniPromise(function (resolve, reject) {
          self._handlers.push({
            onFulfilled: onFulfilled,
            onRejected: onRejected,
            resolve: resolve,
            reject: reject,
          });
          if (self._state !== 'pending') self._flush();
        });
      };
      MiniPromise.prototype.catch = function (onRejected) {
        return this.then(null, onRejected);
      };
      MiniPromise.resolve = function (v) {
        return new MiniPromise(function (resolve) {
          resolve(v);
        });
      };
      MiniPromise.reject = function (e) {
        return new MiniPromise(function (_, reject) {
          reject(e);
        });
      };
      MiniPromise.all = function (arr) {
        arr = arr || [];
        return new MiniPromise(function (resolve, reject) {
          if (!arr.length) return resolve([]);
          var out = new Array(arr.length);
          var left = arr.length;
          for (var i = 0; i < arr.length; i++) {
            (function (idx) {
              MiniPromise.resolve(arr[idx]).then(function (v) {
                out[idx] = v;
                left -= 1;
                if (left === 0) resolve(out);
              }, reject);
            })(i);
          }
        });
      };
      window.Promise = MiniPromise;
    })();
  }

  var ROLES = [
    { id: 'author', label: '作者' },
    { id: 'reviewer', label: '审核' },
    { id: 'approver', label: '批准' },
    { id: 'executor', label: '执行人员' },
    { id: 'reviewer_tail', label: '审核人员（文末等）' },
  ];

  var APP_PREFIX =
    typeof window !== 'undefined' && window.__APP_ROOT__
      ? String(window.__APP_ROOT__).replace(/\/+$/, '')
      : '';

  function apiUrl(path) {
    var p = path.charAt(0) === '/' ? path : '/' + path;
    return APP_PREFIX + p;
  }

  var DEFAULT_FETCH_TIMEOUT_MS = 45000;
  var SIGNERS_LIST_FETCH_TIMEOUT_MS = 90000;
  var SIGN_FILES_LIST_FETCH_TIMEOUT_MS = 90000;
  var HANDOFF_CLAIM_TIMEOUT_MS = 300000;
  var PIECE_BATCH_UPLOAD_TIMEOUT_MS = 120000;
  var PIECE_BATCH_ITEM_TIMEOUT_MS = 90000;
  /** 含压缩包时上传+服务端解压可能较久 */
  var SIGN_ARCHIVE_UPLOAD_TIMEOUT_MS = 600000;
  var SIGN_MULTI_UPLOAD_TIMEOUT_MS = 180000;
  // 文档识别超时（默认 12 小时）。在「系统设置」中按
  // SIGN_DETECT_TIMEOUT_MS 配置；本变量启动后会由 /api/sign/runtime-config 覆盖。
  // 文件量大时识别耗时可能很久，因此默认给长超时并允许在系统设置中调整。
  var DETECT_FETCH_TIMEOUT_MS = 43200000;
  function _detectTimeoutMs() {
    var n = parseInt(window.__signRuntimeConfig &&
      window.__signRuntimeConfig.sign_detect_timeout_ms, 10);
    if (!isFinite(n) || n < 30000) return DETECT_FETCH_TIMEOUT_MS;
    if (n > 86400000) return 86400000;
    return n;
  }
  /** 批量时每文件识别上限，避免单个 SRS 等大文档拖死整批进度 */
  function _batchDetectPerFileTimeoutMs() {
    var full = _detectTimeoutMs();
    var cap = 600000;
    return full > cap ? cap : full;
  }
  function _withBatchDetectTimeout(detectPromise, row) {
    var limitMs = _batchDetectPerFileTimeoutMs();
    var tick = null;
    var stallTimer = null;
    var stall = new Promise(function (_, reject) {
      var started = Date.now();
      tick = setInterval(function () {
        var elapsed = Date.now() - started;
        if (!row) return;
        if (elapsed > 45000 && (row.status === '识别中…' || row.status === '识别重试中…')) {
          row.status = '识别较慢…（' + Math.round(elapsed / 1000) + 's）';
          row.slotTags = ['识别中'];
          try {
            renderBatchWorkbenchTable();
          } catch (_) {}
        }
      }, 15000);
      stallTimer = setTimeout(function () {
        reject(
          new Error(
            '识别超时（单文件超过 ' + Math.round(limitMs / 1000) + ' 秒，常见于软件需求规范等长文档）'
          )
        );
      }, limitMs);
    });
    return Promise.race([detectPromise, stall]).finally(function () {
      if (tick) clearInterval(tick);
      if (stallTimer) clearTimeout(stallTimer);
    });
  }
  function _detectRetryTimes() {
    var n = parseInt(window.__signRuntimeConfig &&
      window.__signRuntimeConfig.sign_detect_retry_times, 10);
    if (!isFinite(n) || n < 0) return 1;
    if (n > 3) return 3;
    return n;
  }
  function refreshRuntimeConfig() {
    return fetchJson(apiUrl('/api/sign/runtime-config'), { timeoutMs: 15000 })
      .then(function (r) {
        var d = (r && r.data) || {};
        if (d && d.ok !== false) {
          window.__signRuntimeConfig = {
            sign_detect_timeout_ms: d.sign_detect_timeout_ms,
            sign_detect_retry_times: d.sign_detect_retry_times,
          };
          // 同步给老变量，老入口（手动重新识别等）也能用上
          if (typeof d.sign_detect_retry_times === 'number') {
            __aiwordHandoffDetectRetries = Math.max(0, Math.min(3, d.sign_detect_retry_times));
          }
        }
      })
      .catch(function () { /* 静默：失败用默认值 */ });
  }
  var __refreshSignersPromise = null;
  var __refreshSignersDeferTimer = null;
  var __refreshFileListPromise = null;
  var __refreshFileListDeferTimer = null;

  function _fetchTimeoutError(timeoutMs) {
    var sec = Math.max(1, Math.round((timeoutMs || DEFAULT_FETCH_TIMEOUT_MS) / 1000));
    return new Error(
      '请求超时（' + sec + 's）：远程 MySQL 或上传数据较大时可能较慢，将自动重试；请确认 aiprintword 已启动且网络可达'
    );
  }

  function _isUploadCancelledError(err) {
    return !!(err && err.uploadCancelled);
  }

  function _uploadCancelledError() {
    var e = new Error('已取消上传');
    e.uploadCancelled = true;
    return e;
  }

  function _isRetriableFetchError(err) {
    if (_isUploadCancelledError(err)) return false;
    var msg = String((err && err.message) || err || '');
    if (/已取消上传/.test(msg)) return false;
    return /超时|timeout|网络|failed to fetch|network|连接|reset|ECONN/i.test(msg);
  }

  function _fetchTextCompat(url, options) {
    var opts = Object.assign({}, options || {});
    var timeoutMs =
      typeof opts.timeoutMs === 'number' && opts.timeoutMs > 0
        ? opts.timeoutMs
        : DEFAULT_FETCH_TIMEOUT_MS;
    var externalSignal = opts.signal;
    delete opts.timeoutMs;
    delete opts.signal;
    if (typeof AbortController !== 'undefined' && typeof fetch === 'function') {
      var ac = new AbortController();
      var timedOut = false;
      var userAborted = false;
      if (externalSignal) {
        if (externalSignal.aborted) {
          return Promise.reject(_uploadCancelledError());
        }
        externalSignal.addEventListener(
          'abort',
          function () {
            userAborted = true;
            try {
              ac.abort();
            } catch (_) {}
          },
          { once: true }
        );
      }
      var tid = setTimeout(function () {
        timedOut = true;
        try {
          ac.abort();
        } catch (_) {}
      }, timeoutMs);
      var fetchOpts = Object.assign({ credentials: 'include' }, opts, { signal: ac.signal });
      return fetch(url, fetchOpts)
        .then(function (res) {
          clearTimeout(tid);
          return res.text().then(function (text) {
            return { res: res, text: text };
          });
        })
        .catch(function (err) {
          clearTimeout(tid);
          if (userAborted || (externalSignal && externalSignal.aborted)) {
            throw _uploadCancelledError();
          }
          if (timedOut || (err && (err.name === 'AbortError' || /aborted/i.test(String(err.message || ''))))) {
            throw _fetchTimeoutError(timeoutMs);
          }
          throw err;
        });
    }
    return new Promise(function (resolve, reject) {
      var xhr = new XMLHttpRequest();
      xhr.open((opts.method || 'GET').toUpperCase(), url, true);
      xhr.withCredentials = true;
      try {
        if (opts.headers) {
          Object.keys(opts.headers).forEach(function (k) {
            xhr.setRequestHeader(k, opts.headers[k]);
          });
        }
      } catch (_) {}
      xhr.onreadystatechange = function () {
        if (xhr.readyState !== 4) return;
        var resLike = {
          status: xhr.status || 0,
          ok: xhr.status >= 200 && xhr.status < 300,
        };
        resolve({ res: resLike, text: xhr.responseText || '' });
      };
      xhr.onerror = function () {
        reject(new Error('网络请求失败（浏览器兼容模式）'));
      };
      xhr.timeout = timeoutMs;
      xhr.ontimeout = function () {
        reject(_fetchTimeoutError(timeoutMs));
      };
      xhr.send(opts.body || null);
    });
  }

  function fetchJson(url, options) {
    var opts = Object.assign({ credentials: 'include' }, options || {});
    return _fetchTextCompat(url, opts).then(function (pack) {
      var res = pack.res;
      var text = pack.text || '';
        var t = text.trim();
        if (t.charAt(0) === '<') {
          var m = (opts && opts.method ? String(opts.method).toUpperCase() : 'GET');
          var hint =
            res.status === 404
              ? '接口不存在（请确认已保存最新代码并重启 python app.py）。'
              : res.status === 405
                ? '接口方法不被允许（HTTP 405）。通常是后端未重启或你访问的地址不是最新的服务进程；请重启 python app.py，并用 http://127.0.0.1:5050/sign 打开后 Ctrl+F5 强制刷新。'
              : '服务器返回了网页而不是 JSON。';
          throw new Error(hint + ' ' + m + ' ' + url + ' HTTP ' + res.status);
        }
        try {
          return { res: res, data: JSON.parse(text) };
        } catch (e) {
          throw new Error('接口返回无法解析为 JSON：' + t.slice(0, 160));
        }
    });
  }

  function fetchJsonWithRetry(url, options, retryOpt) {
    retryOpt = retryOpt || {};
    var maxTry = Math.max(1, retryOpt.maxTry || 3);
    var delayMs = typeof retryOpt.delayMs === 'number' ? retryOpt.delayMs : 1000;
    var onRetry = retryOpt.onRetry;
    function attempt(n) {
      return fetchJson(url, options).catch(function (err) {
        if (n >= maxTry || !_isRetriableFetchError(err)) {
          throw err;
        }
        if (typeof onRetry === 'function') {
          try {
            onRetry(n + 1, maxTry, err);
          } catch (_) {}
        }
        return new Promise(function (resolve) {
          setTimeout(resolve, delayMs * n);
        }).then(function () {
          return attempt(n + 1);
        });
      });
    }
    return attempt(1);
  }

  var __pageProgressDepth = 0;
  var __pageProgressTicker = null;
  var __pageProgressStartedAt = 0;
  var __pageProgressBaseText = '';
  var __pageProgressDone = 0;
  var __pageProgressTotal = 0;

  function ensureSignPageProgressDom() {
    var wrap = document.getElementById('signPageProgressWrap');
    if (wrap) return wrap;
    var st = document.getElementById('signPageProgressStyle');
    if (!st) {
      st = document.createElement('style');
      st.id = 'signPageProgressStyle';
      st.textContent =
        '#signPageProgressWrap{display:none;position:sticky;top:0;z-index:1200;margin:0;padding:8px 12px;background:#e8f0fe;border-bottom:1px solid #90caf9;box-shadow:0 2px 8px rgba(15,23,42,.08)}' +
        '#signPageProgressWrap.is-active{display:block}' +
        '#signPageProgressBar{width:100%;height:12px;accent-color:var(--accent,#1a73e8);margin:0}' +
        '#signPageProgressText{margin-top:6px;font-size:.88rem;color:#0d47a1;line-height:1.45;white-space:pre-wrap}';
      document.head.appendChild(st);
    }
    wrap = document.createElement('div');
    wrap.id = 'signPageProgressWrap';
    wrap.setAttribute('role', 'status');
    wrap.setAttribute('aria-live', 'polite');
    wrap.innerHTML =
      '<progress id="signPageProgressBar" max="100" value="0"></progress>' +
      '<div id="signPageProgressText"></div>';
    var anchor = document.getElementById('signBootstrapBanner');
    if (anchor && anchor.parentNode) {
      anchor.parentNode.insertBefore(wrap, anchor.nextSibling);
    } else {
      document.body.insertBefore(wrap, document.body.firstChild);
    }
    return wrap;
  }

  function _renderSignPageProgress() {
    ensureSignPageProgressDom();
    var wrap = document.getElementById('signPageProgressWrap');
    var bar = document.getElementById('signPageProgressBar');
    var txt = document.getElementById('signPageProgressText');
    if (!wrap || !bar || !txt) return;
    if (__pageProgressDepth < 1) {
      wrap.classList.remove('is-active');
      wrap.style.display = 'none';
      bar.value = 0;
      txt.textContent = '';
      return;
    }
    wrap.classList.add('is-active');
    wrap.style.display = 'block';
    var elapsed = __pageProgressStartedAt
      ? Math.max(0, Math.round((Date.now() - __pageProgressStartedAt) / 1000))
      : 0;
    var total = Math.max(0, parseInt(__pageProgressTotal, 10) || 0);
    var done = Math.max(0, parseInt(__pageProgressDone, 10) || 0);
    if (total > 0 && done > total) done = total;
    var pct;
    var line;
    if (total > 0) {
      pct = Math.round((done / total) * 100);
      bar.max = total;
      bar.value = done;
      line =
        (__pageProgressBaseText || '处理中') +
        ' ' +
        done +
        '/' +
        total +
        '（' +
        pct +
        '%）';
    } else {
      pct = Math.min(92, 6 + Math.floor(elapsed / 2));
      bar.max = 100;
      bar.value = pct;
      line = (__pageProgressBaseText || '处理中…') + '（已等待 ' + elapsed + ' 秒）';
    }
    txt.textContent = line;
  }

  function beginPageProgress(text, opt) {
    opt = opt || {};
    ensureSignPageProgressDom();
    __pageProgressDepth++;
    if (__pageProgressDepth === 1) {
      __pageProgressStartedAt = Date.now();
      __pageProgressDone = 0;
      __pageProgressTotal = 0;
    }
    if (text) __pageProgressBaseText = String(text);
    if (typeof opt.done === 'number' && typeof opt.total === 'number') {
      __pageProgressDone = opt.done;
      __pageProgressTotal = opt.total;
    }
    _renderSignPageProgress();
    if (!__pageProgressTicker) {
      __pageProgressTicker = setInterval(_renderSignPageProgress, 1000);
    }
  }

  function updatePageProgress(text, opt) {
    opt = opt || {};
    if (text) __pageProgressBaseText = String(text);
    if (typeof opt.done === 'number') __pageProgressDone = opt.done;
    if (typeof opt.total === 'number') __pageProgressTotal = opt.total;
    if (__pageProgressDepth > 0) _renderSignPageProgress();
  }

  function endPageProgress() {
    if (__pageProgressDepth > 0) __pageProgressDepth--;
    if (__pageProgressDepth > 0) {
      _renderSignPageProgress();
      return;
    }
    if (__pageProgressTicker) {
      clearInterval(__pageProgressTicker);
      __pageProgressTicker = null;
    }
    __pageProgressStartedAt = 0;
    __pageProgressBaseText = '';
    __pageProgressDone = 0;
    __pageProgressTotal = 0;
    _renderSignPageProgress();
  }

  function runWithPageProgress(text, fn, opt) {
    opt = opt || {};
    if (opt.skipPageProgress) return Promise.resolve().then(fn);
    beginPageProgress(text, opt);
    return Promise.resolve()
      .then(fn)
      .finally(function () {
        endPageProgress();
      });
  }

  /**
   * 异步操作期间禁用按钮并显示 spinner，避免用户以为未点击。
   * @param {Object} opt 若 opt.skipRestoreDisabled，结束时不再恢复 disabled（由业务在 finally 里设置）
   * @param {Object} opt skipPageProgress 为 true 时不显示顶栏进度（如已有专用进度区且避免重复）
   */
  function withButtonBusy(btn, busyLabel, fn, opt) {
    opt = opt || {};
    var pageLabel = opt.pageProgressLabel || busyLabel || '处理中…';
    if (!opt.skipPageProgress) beginPageProgress(pageLabel);
    function wrapFn() {
      if (!btn) return Promise.resolve(fn());
      if (btn.getAttribute('aria-busy') === 'true') return Promise.resolve();
      var prevHtml = btn.innerHTML;
      var wasDisabled = btn.disabled;
      btn.setAttribute('aria-busy', 'true');
      btn.disabled = true;
      btn.innerHTML =
        '<span class="spinner" aria-hidden="true"></span> ' + (busyLabel || '处理中…');
      function restoreBtn() {
        btn.removeAttribute('aria-busy');
        btn.innerHTML = prevHtml;
        if (!opt.skipRestoreDisabled) {
          btn.disabled = wasDisabled;
        }
      }
      return Promise.resolve().then(fn).then(restoreBtn, restoreBtn);
    }
    return wrapFn().finally(function () {
      if (!opt.skipPageProgress) endPageProgress();
    });
  }

  if (typeof window !== 'undefined' && window.location.protocol === 'file:') {
    var w = document.getElementById('fileProtoWarn');
    if (w) w.style.display = 'block';
  }
  if (typeof window !== 'undefined') {
    window.addEventListener('error', function (ev) {
      try {
        var b = document.getElementById('signBootstrapBanner');
        if (!b) return;
        b.style.display = 'block';
        var msg = (ev && ev.message) ? ev.message : '未知错误';
        var src = (ev && ev.filename) ? String(ev.filename).split('/').pop() : '';
        var ln = (ev && ev.lineno) ? String(ev.lineno) : '';
        b.textContent = '前端脚本异常：' + msg + (src ? ' @' + src : '') + (ln ? ':' + ln : '');
        try {
          var m = document.getElementById('aiwordHandoffLoadingMask');
          if (m) m.classList.remove('show');
        } catch (_) {}
      } catch (_) {}
    });
    window.addEventListener('unhandledrejection', function (ev) {
      try {
        var b = document.getElementById('signBootstrapBanner');
        if (!b) return;
        b.style.display = 'block';
        var reason = (ev && ev.reason) ? String(ev.reason.message || ev.reason) : 'Promise 未处理异常';
        b.textContent = '前端脚本异常：' + reason;
        try {
          var m = document.getElementById('aiwordHandoffLoadingMask');
          if (m) m.classList.remove('show');
        } catch (_) {}
      } catch (_) {}
    });
  }

  var fileInput = document.getElementById('fileInput');
  var dirInput = document.getElementById('dirInput');
  var fileHint = document.getElementById('fileHint');
  var saveBtn = document.getElementById('saveBtn');
  var fileListEl = document.getElementById('fileList');
  var listHint = document.getElementById('listHint');
  var fileListActionMsg = document.getElementById('fileListActionMsg');
  var roleChecks = document.getElementById('roleChecks');
  var rolePanels = document.getElementById('rolePanels');
  var submitBtn = document.getElementById('submitBtn');
  var errMsg = document.getElementById('errMsg');
  var signedListEl = document.getElementById('signedList');
  var signedHint = document.getElementById('signedHint');
  var signedListActionMsg = document.getElementById('signedListActionMsg');
  var signerLibHint = document.getElementById('signerLibHint');
  var newSignerName = document.getElementById('newSignerName');
  var addSignerBtn = document.getElementById('addSignerBtn');
  var signerListEl = document.getElementById('signerListEl');
  var signerPagerInfo = document.getElementById('signerPagerInfo');
  var signerPrevBtn = document.getElementById('signerPrevBtn');
  var signerNextBtn = document.getElementById('signerNextBtn');
  var needSignTable = document.getElementById('needSignTable');
  var redetectRolesBtn = document.getElementById('redetectRolesBtn');
  var batchApplyRoleMapBtn = document.getElementById('batchApplyRoleMapBtn');
  var saveRoleConfigBtn = document.getElementById('saveRoleConfigBtn');
  var batchSelectAll = document.getElementById('batchSelectAll');
  var batchModeCb = document.getElementById('batchModeCb');
  var batchResultMsg = document.getElementById('batchResultMsg');
  var aiwordBatchStrategyCard = document.getElementById('aiwordBatchStrategyCard');
  var aiwordBatchGroupStrategy = document.getElementById('aiwordBatchGroupStrategy');
  var aiwordBatchExecMode = document.getElementById('aiwordBatchExecMode');
  var aiwordBatchPreviewBtn = document.getElementById('aiwordBatchPreviewBtn');
  var aiwordBatchApplySelectBtn = document.getElementById('aiwordBatchApplySelectBtn');
  var aiwordBatchExportBtn = document.getElementById('aiwordBatchExportBtn');
  var aiwordBatchRunBtn = document.getElementById('aiwordBatchRunBtn');
  var aiwordBatchStrategyMsg = document.getElementById('aiwordBatchStrategyMsg');
  var aiwordBatchGroupList = document.getElementById('aiwordBatchGroupList');
  var batchWorkbenchCard = document.getElementById('batchWorkbenchCard');
  var batchWorkbenchBody = document.getElementById('batchWorkbenchBody');
  var batchWorkbenchMsg = document.getElementById('batchWorkbenchMsg');
  var batchWorkbenchProgressWrap = document.getElementById('batchWorkbenchProgressWrap');
  var batchWorkbenchProgressBar = document.getElementById('batchWorkbenchProgressBar');
  var batchWorkbenchProgressText = document.getElementById('batchWorkbenchProgressText');
  var batchWorkbenchHiddenPicks = document.getElementById('batchWorkbenchHiddenPicks');
  var batchWorkbenchSelectAll = document.getElementById('batchWorkbenchSelectAll');
  var batchWorkbenchHeadCheck = document.getElementById('batchWorkbenchHeadCheck');
  // sign_file.html 以工作台为主，无 legacy「batchSelectAll」；复用工作台全选节点
  if (!batchSelectAll && batchWorkbenchSelectAll) {
    batchSelectAll = batchWorkbenchSelectAll;
  }
  var batchWorkbenchDetectBtn = document.getElementById('batchWorkbenchDetectBtn');
  var batchWorkbenchApplyBtn = document.getElementById('batchWorkbenchApplyBtn');
  var batchWorkbenchDeleteBtn = document.getElementById('batchWorkbenchDeleteBtn');
  var batchWorkbenchLocaleBulk = document.getElementById('batchWorkbenchLocaleBulk');
  var batchWorkbenchLocaleApplyBtn = document.getElementById('batchWorkbenchLocaleApplyBtn');
  var batchWorkbenchExportIssuesBtn = document.getElementById('batchWorkbenchExportIssuesBtn');
  var batchWorkbenchSignBtn = document.getElementById('batchWorkbenchSignBtn');
  var batchWorkbenchRefreshBtn = document.getElementById('batchWorkbenchRefreshBtn');
  var batchWorkbenchSelectFilteredBtn = document.getElementById('batchWorkbenchSelectFilteredBtn');
  var batchWorkbenchFilterName = document.getElementById('batchWorkbenchFilterName');
  var batchWorkbenchFilterClearBtn = document.getElementById('batchWorkbenchFilterClearBtn');
  var batchWorkbenchFilterNotSignableBtn = document.getElementById('batchWorkbenchFilterNotSignableBtn');
  var batchWorkbenchSelectByStatusBtn = document.getElementById('batchWorkbenchSelectByStatusBtn');
  var batchWorkbenchFilterNameTags = document.getElementById('batchWorkbenchFilterNameTags');
  var batchWorkbenchFilterStatusWrap = document.getElementById('batchWorkbenchFilterStatusWrap');
  var batchWorkbenchFilterStatusToggle = document.getElementById(
    'batchWorkbenchFilterStatusToggle'
  );
  var batchWorkbenchFilterStatusBox = document.getElementById('batchWorkbenchFilterStatusBox');
  var batchWorkbenchFilterSlotWrap = document.getElementById('batchWorkbenchFilterSlotWrap');
  var batchWorkbenchFilterSlotToggle = document.getElementById('batchWorkbenchFilterSlotToggle');
  var batchWorkbenchFilterSlotBox = document.getElementById('batchWorkbenchFilterSlotBox');
  var batchWorkbenchStatusSelectAllBtn = document.getElementById(
    'batchWorkbenchStatusSelectAllBtn'
  );
  var batchWorkbenchStatusClearBtn = document.getElementById('batchWorkbenchStatusClearBtn');
  var batchWorkbenchSlotSelectAllBtn = document.getElementById('batchWorkbenchSlotSelectAllBtn');
  var batchWorkbenchSlotClearBtn = document.getElementById('batchWorkbenchSlotClearBtn');
  var batchWorkbenchFilterHint = document.getElementById('batchWorkbenchFilterHint');
  var batchWorkbenchAdvancedCb = document.getElementById('batchWorkbenchAdvancedCb');
  var needSignActionMsg = document.getElementById('needSignActionMsg');
  var saveUploadFeedback = document.getElementById('saveUploadFeedback');
  var saveUploadProgressWrap = document.getElementById('saveUploadProgressWrap');
  var saveUploadProgressBar = document.getElementById('saveUploadProgressBar');
  var saveUploadProgressText = document.getElementById('saveUploadProgressText');
  var signSourceMode = document.getElementById('signSourceMode');
  var libraryRolesModeRow = document.getElementById('libraryRolesModeRow');
  var libraryRolesUseChecksCb = document.getElementById('libraryRolesUseChecksCb');
  var signerErrMsg = document.getElementById('signerErrMsg');
  var libSignerSelect = document.getElementById('libSignerSelect');
  var libSignerFilter = document.getElementById('libSignerFilter');
  var libStrokeSetSelect = document.getElementById('libStrokeSetSelect');
  var libLocaleSelect = document.getElementById('libLocaleSelect');
  var libClearSigBtn = document.getElementById('libClearSigBtn');
  var libClearDateBtn = document.getElementById('libClearDateBtn');
  var libLoadStrokesBtn = document.getElementById('libLoadStrokesBtn');
  var libSaveStrokesBtn = document.getElementById('libSaveStrokesBtn');
  var libSaveSigOnlyBtn = document.getElementById('libSaveSigOnlyBtn');
  var libSaveDateOnlyBtn = document.getElementById('libSaveDateOnlyBtn');
  var libSaveStrokeSetBtn = document.getElementById('libSaveStrokeSetBtn');
  var libStrokeFeedback = document.getElementById('libStrokeFeedback');
  var strokeItemsHint = document.getElementById('strokeItemsHint');
  var strokeItemCatSelect = document.getElementById('strokeItemCatSelect');
  var strokeItemsActionMsg = document.getElementById('strokeItemsActionMsg');
  var strokeItemSearchInput = document.getElementById('strokeItemSearchInput');
  // 页面拆分：文件签名页 / 素材录入页（两页复用同一脚本；按 DOM 存在与否决定绑定哪些交互）
  var pageFileSign = document.getElementById('pageFileSign');
  var pageMaterials = document.getElementById('pageMaterials');
  var IS_FILE_SIGN_PAGE = !!pageFileSign;
  var IS_MATERIALS_PAGE = !!pageMaterials;
  var strokeItemSearchBtn = document.getElementById('strokeItemSearchBtn');
  var strokeItemPager = document.getElementById('strokeItemPager');
  var strokeItemPagerInfo = document.getElementById('strokeItemPagerInfo');
  var strokeItemPrevBtn = document.getElementById('strokeItemPrevBtn');
  var strokeItemNextBtn = document.getElementById('strokeItemNextBtn');
  var strokeItemListEl = document.getElementById('strokeItemListEl');
  var signedSearchRow = document.getElementById('signedSearchRow');
  var signedSearchInput = document.getElementById('signedSearchInput');
  var signedSearchBtn = document.getElementById('signedSearchBtn');
  var signedPagerInfo = document.getElementById('signedPagerInfo');
  var signedPrevBtn = document.getElementById('signedPrevBtn');
  var signedNextBtn = document.getElementById('signedNextBtn');

  var btnRefreshSigners = document.getElementById('btnRefreshSigners');
  var btnRefreshStrokeItems = document.getElementById('btnRefreshStrokeItems');
  var btnRefreshFiles = document.getElementById('btnRefreshFiles');
  var batchDeleteFilesBtn = document.getElementById('batchDeleteFilesBtn');
  var btnRefreshSigned = document.getElementById('btnRefreshSigned');
  var pieceDateCard = document.getElementById('pieceDateCard');
  var pieceClearBtn = document.getElementById('pieceClearBtn');
  var pieceCanvasEl = document.getElementById('piece_canvas');
  var pieceHint = document.getElementById('pieceHint');
  var currentSignerName = document.getElementById('currentSignerName');
  var pieceCurrentSignerName = document.getElementById('pieceCurrentSignerName');
  var pieceDigitChecks = document.getElementById('pieceDigitChecks');
  var pieceMonthChecks = document.getElementById('pieceMonthChecks');
  var pieceBatchStartBtn = document.getElementById('pieceBatchStartBtn');
  var pieceBatchNextBtn = document.getElementById('pieceBatchNextBtn');
  var pieceBatchUploadBtn = document.getElementById('pieceBatchUploadBtn');
  var pieceBatchCancelBtn = document.getElementById('pieceBatchCancelBtn');
  var pieceBatchStatus = document.getElementById('pieceBatchStatus');

  // 页面拆分后：仅校验“当前页面需要的 DOM”。避免提示“请用完整 sign.html 覆盖”。
  var _signDomRequired = [];
  function _req(id, el) {
    _signDomRequired.push([id, el]);
  }
  // 两页共用（脚本运行与告警用）
  _req('signBootstrapBanner', document.getElementById('signBootstrapBanner'));
  if (IS_FILE_SIGN_PAGE) {
    _req('fileInput', fileInput);
    _req('fileHint', fileHint);
    _req('saveBtn', saveBtn);
    _req('fileListEl', fileListEl);
    _req('listHint', listHint);
    _req('roleChecks', roleChecks);
    _req('rolePanels', rolePanels);
    _req('submitBtn', submitBtn);
    _req('errMsg', errMsg);
    _req('signedListEl', signedListEl);
    _req('signedHint', signedHint);
    _req('needSignTable', needSignTable);
    _req('redetectRolesBtn', redetectRolesBtn);
    _req('batchSelectAll', batchSelectAll);
    _req('batchModeCb', batchModeCb);
    _req('batchResultMsg', batchResultMsg);
    _req('needSignActionMsg', needSignActionMsg);
    _req('saveUploadFeedback', saveUploadFeedback);
    _req('signSourceMode', signSourceMode);
    _req('signedSearchRow', signedSearchRow);
    _req('signedSearchInput', signedSearchInput);
    _req('signedSearchBtn', signedSearchBtn);
    _req('signedPagerInfo', signedPagerInfo);
    _req('signedPrevBtn', signedPrevBtn);
    _req('signedNextBtn', signedNextBtn);
  }
  if (IS_MATERIALS_PAGE) {
    _req('signerLibHint', signerLibHint);
    _req('newSignerName', newSignerName);
    _req('addSignerBtn', addSignerBtn);
    _req('signerListEl', signerListEl);
    _req('signerPagerInfo', signerPagerInfo);
    _req('signerPrevBtn', signerPrevBtn);
    _req('signerNextBtn', signerNextBtn);
    _req('signerErrMsg', signerErrMsg);
    _req('libSignerSelect', libSignerSelect);
    _req('libSignerFilter', libSignerFilter);
    _req('libStrokeSetSelect', libStrokeSetSelect);
    _req('libLocaleSelect', libLocaleSelect);
    _req('libClearSigBtn', libClearSigBtn);
    _req('libClearDateBtn', libClearDateBtn);
    _req('libLoadStrokesBtn', libLoadStrokesBtn);
    _req('libSaveStrokesBtn', libSaveStrokesBtn);
    _req('strokeItemsHint', strokeItemsHint);
    _req('strokeItemCatSelect', strokeItemCatSelect);
    _req('strokeItemSearchInput', strokeItemSearchInput);
    _req('strokeItemSearchBtn', strokeItemSearchBtn);
    _req('strokeItemPagerInfo', strokeItemPagerInfo);
    _req('strokeItemPrevBtn', strokeItemPrevBtn);
    _req('strokeItemNextBtn', strokeItemNextBtn);
    _req('strokeItemListEl', strokeItemListEl);
    _req('pieceDateCard', pieceDateCard);
    _req('pieceClearBtn', pieceClearBtn);
    _req('pieceCanvasEl', pieceCanvasEl);
    _req('pieceHint', pieceHint);
    _req('currentSignerName', currentSignerName);
    _req('pieceCurrentSignerName', pieceCurrentSignerName);
    _req('pieceDigitChecks', pieceDigitChecks);
    _req('pieceMonthChecks', pieceMonthChecks);
    _req('pieceBatchStartBtn', pieceBatchStartBtn);
    _req('pieceBatchNextBtn', pieceBatchNextBtn);
    _req('pieceBatchUploadBtn', pieceBatchUploadBtn);
    _req('pieceBatchCancelBtn', pieceBatchCancelBtn);
    _req('pieceBatchStatus', pieceBatchStatus);
  }
  var _signDomMissing = [];
  for (var _di = 0; _di < _signDomRequired.length; _di++) {
    if (!_signDomRequired[_di][1]) _signDomMissing.push(_signDomRequired[_di][0]);
  }
  if (_signDomMissing.length) {
    var _failMsg =
      '签名脚本未启动：页面缺少 id 为 ' +
      _signDomMissing.slice(0, 14).join(', ') +
      (_signDomMissing.length > 14 ? '…（共 ' + _signDomMissing.length + ' 项）' : ' 的节点');
    window.__SIGN_PAGE_BOOT_FAIL_MSG = _failMsg;
    try {
      window.__SIGN_PAGE_BOOT_HALTED = true;
    } catch (_) {}
    try {
      var _ban = document.getElementById('signBootstrapBanner');
      if (_ban) {
        _ban.style.display = 'block';
        _ban.textContent = _failMsg + '。请 Ctrl+F5 强刷；确认页面文件与 static/js/sign-page.js 已同步部署，并已重启 python app.py。';
      }
      var ph = document.getElementById('signerLibHint');
      if (ph) ph.textContent = _failMsg;
      var lh = document.getElementById('listHint');
      if (lh) lh.textContent = _failMsg;
    } catch (_) {}
    return;
  }

  /** 清除「需签角色」表内行反馈、重新识别旁、上传保存旁等非「生成文档」区提示 */
  function clearNeedSignScopedFeedbacks() {
    try {
      if (needSignActionMsg) {
        needSignActionMsg.style.display = 'none';
        needSignActionMsg.textContent = '';
        needSignActionMsg.className = 'btn-inline-feedback is-error';
      }
      if (saveUploadFeedback) {
        saveUploadFeedback.style.display = 'none';
        saveUploadFeedback.textContent = '';
        saveUploadFeedback.className = 'btn-inline-feedback is-error';
      }
      if (needSignTable) {
        var nodes = needSignTable.querySelectorAll('[id^="needSignRowMsg_"]');
        for (var ni = 0; ni < nodes.length; ni++) {
          nodes[ni].style.display = 'none';
          nodes[ni].textContent = '';
          nodes[ni].className = 'btn-inline-feedback';
        }
      }
    } catch (_) {}
  }

  function setNeedSignActionFeedback(s) {
    if (!needSignActionMsg) return;
    if (!s) {
      needSignActionMsg.style.display = 'none';
      needSignActionMsg.textContent = '';
      needSignActionMsg.className = 'btn-inline-feedback is-error';
      return;
    }
    clearFileRegionErr();
    if (signerErrMsg) {
      signerErrMsg.style.display = 'none';
      signerErrMsg.textContent = '';
    }
    needSignActionMsg.style.display = 'block';
    needSignActionMsg.textContent = s;
    needSignActionMsg.className = 'btn-inline-feedback is-error';
  }

  function setSaveUploadFeedback(s, tone) {
    if (!saveUploadFeedback) return;
    if (!s) {
      saveUploadFeedback.style.display = 'none';
      saveUploadFeedback.textContent = '';
      saveUploadFeedback.className = 'btn-inline-feedback';
      return;
    }
    clearFileRegionErr();
    if (signerErrMsg) {
      signerErrMsg.style.display = 'none';
      signerErrMsg.textContent = '';
    }
    var cls = 'btn-inline-feedback';
    if (tone === 'ok') cls += ' is-ok';
    else if (tone === 'warn') cls += ' is-warn';
    else if (tone !== 'info') cls += ' is-error';
    saveUploadFeedback.style.display = 'block';
    saveUploadFeedback.textContent = s;
    saveUploadFeedback.className = cls;
  }

  function setSaveUploadProgress(show, pct, text) {
    if (!saveUploadProgressWrap) return;
    if (!show) {
      saveUploadProgressWrap.style.display = 'none';
      if (saveUploadProgressBar) saveUploadProgressBar.value = 0;
      if (saveUploadProgressText) saveUploadProgressText.textContent = '';
      return;
    }
    saveUploadProgressWrap.style.display = 'block';
    var p = Math.max(0, Math.min(100, Number(pct) || 0));
    if (saveUploadProgressBar) {
      saveUploadProgressBar.value = p;
      saveUploadProgressBar.removeAttribute('value');
      saveUploadProgressBar.setAttribute('value', String(p));
    }
    if (saveUploadProgressText) {
      saveUploadProgressText.textContent = text || '处理中…';
    }
  }

  function pendingSelectionHasArchive() {
    for (var i = 0; i < pendingSignFiles.length; i++) {
      var nm = String((pendingSignFiles[i] && pendingSignFiles[i].name) || '').toLowerCase();
      if (/\.(zip|7z|rar)$/.test(nm)) return true;
    }
    return false;
  }

  function _saveUploadTimeoutMs() {
    if (pendingSelectionHasArchive()) return SIGN_ARCHIVE_UPLOAD_TIMEOUT_MS;
    if (pendingSignFiles.length > 1) return SIGN_MULTI_UPLOAD_TIMEOUT_MS;
    return PIECE_BATCH_UPLOAD_TIMEOUT_MS;
  }

  function setRoleRowFeedback(rid, s) {
    var el = document.getElementById('needSignRowMsg_' + rid);
    if (!el) return;
    if (!s) {
      el.style.display = 'none';
      el.textContent = '';
      el.className = 'btn-inline-feedback';
      el.style.whiteSpace = '';
      return;
    }
    clearFileRegionErr();
    if (signerErrMsg) {
      signerErrMsg.style.display = 'none';
      signerErrMsg.textContent = '';
    }
    el.style.display = 'block';
    el.textContent = s;
    el.className = 'btn-inline-feedback is-error';
    el.style.whiteSpace = String(s).indexOf('\n') >= 0 ? 'pre-wrap' : 'normal';
  }

  function setPanelSaveFeedback(el, s, isErr) {
    if (!el) return;
    if (!s) {
      el.style.display = 'none';
      el.textContent = '';
      el.className = 'btn-inline-feedback';
      return;
    }
    clearFileRegionErr();
    if (signerErrMsg) {
      signerErrMsg.style.display = 'none';
      signerErrMsg.textContent = '';
    }
    el.style.display = 'block';
    el.textContent = s;
    el.className = 'btn-inline-feedback' + (isErr ? ' is-error' : ' is-ok');
  }

  // 页面拆分后：素材录入页没有该按钮
  var redetectRolesBtnDefaultHtml = redetectRolesBtn ? redetectRolesBtn.innerHTML : '重新识别需签字角色';

  var canvases = {};
  var selectedFileId = null;
  var savedFiles = [];
  var pendingSignFiles = [];
  // 每个文件的 UI/检测缓存：避免切换文件清空选择；避免反复自动识别
  var fileUiCache = {}; // fileId -> { detectedOnce, lastDetectData, lastDetectError, checkedRoleIds, roleItemFilterQ, roleSaveSigner, currentRoleMap }
  var lastDetectData = null;
  /** 与 lastDetectData 对应的 file_id，用于避免列表重绘时误判“已检测”而跳过 /api/sign/detect */
  var lastDetectFileId = null;
  /** 最近一次 /api/sign/detect 失败时的错误文案（用于提示，不阻塞手动勾选角色） */
  var lastDetectError = '';
  /** 正在请求检测的 file_id，避免同一文件并发重复 detect */
  var detectInFlightFor = null;
  /** 每次发起 detect 自增，用于 finally 中只恢复「最后一次」请求的 UI 状态 */
  var detectRequestSeq = 0;
  var fileDetectRequestToken = 0;
  /** 交错请求时仅「当前这一轮」结束才清除 detectInFlightFor，避免长时间运行后按钮/状态卡死 */
  var detectEpoch = 0;
  /** aiword 交接透传的编审批/日期（handoff 响应头 JSON），用于首文件识别后预填签署人 */
  var __aiwordHandoffCtx = null;
  var __aiwordHandoffTargetFileId = null;
  /** aiword 首入：推迟一次自动识别，避免列表刚写入时与首次 detect 竞态导致识别不准 */
  var __aiwordDeferDetectFileId = null;
  /** aiword 交接后 detect 失败时的剩余重试次数（由 kickoff 传入）
   *  注意：detect 超时已统一拉到 5 分钟，超时本身就代表后端在跑，
   *  再 retry 只会让用户看到「请求一直刷」。这里 1 次足够覆盖偶发抖动。 */
  var __aiwordHandoffDetectRetries = 1;
  /** aiword 首次自动识别后再做一次“等效手动重识别”，提高稳定性（fileId -> bool）。 */
  var __aiwordHandoffAutoRedetectDoneFor = {};
  /** aiword 批量交接：file_id -> ctx */
  var __aiwordHandoffCtxByFileId = {};
  /** aiword 批量策略分组缓存 */
  var __aiwordBatchGroups = [];
  /** 批量工作台：每文件一行配置（编审批/日期/版本/状态） */
  var __batchWorkbenchRows = {};
  var __wbFilterNameTerms = [];
  var __wbFilterNameDraft = '';
  var __wbFilterStatuses = [];
  var __wbFilterSlotTags = [];
  var __wbAvailableSlotTags = [];
  /** 处理状态筛选项（与 _workbenchStatusBucket 返回值一致），固定全集 */
  var WORKBENCH_STATUS_FILTER_GROUPS = [
    {
      title: '识别异常',
      items: [
        { value: '识别有误', label: '识别有误（人工登记）' },
        { value: '识别失败', label: '识别失败' },
        { value: '识别超时', label: '识别超时' },
        { value: '未识别', label: '未识别（含未识别到签字位）' },
        { value: '签字位不完整', label: '签字位不完整（含缺姓名/日期位）' },
      ],
    },
    {
      title: '素材与就绪',
      items: [
        { value: '待匹配', label: '待匹配' },
        { value: '就绪', label: '就绪' },
        { value: '部分就绪', label: '部分就绪' },
        { value: '无可用素材', label: '无可用素材' },
      ],
    },
    {
      title: '流水线进行中',
      items: [
        { value: '识别中', label: '识别中/匹配素材/识别较慢/重试中' },
        { value: '签字中', label: '签字中…' },
      ],
    },
    {
      title: '签字结果',
      items: [
        { value: '已签字', label: '已签字' },
        { value: '签字失败', label: '签字失败' },
        { value: '保存映射失败', label: '保存映射失败' },
      ],
    },
    {
      title: '其它',
      items: [
        { value: '无需签字', label: '无需签字' },
        { value: '待处理', label: '待处理（— / 未跑流水线）' },
      ],
    },
  ];
  /** 签字位筛选：固定全集（不依赖当前列表是否出现过） */
  var WORKBENCH_SLOT_FILTER_GROUPS = [
    {
      title: '结论',
      items: [
        { value: '可签', label: '可签' },
        { value: '不可签', label: '不可签' },
        { value: '位齐全', label: '位齐全' },
        { value: '待确认', label: '待确认' },
        { value: '无需签字', label: '无需签字' },
        { value: '缺姓名位', label: '缺姓名位' },
        { value: '缺日期位', label: '缺日期位' },
        { value: '缺姓名+日期位', label: '缺姓名+日期位' },
      ],
    },
    {
      title: '缺位标签',
      items: [
        { value: '姓名位缺失', label: '姓名位缺失' },
        { value: '日期位缺失', label: '日期位缺失' },
      ],
    },
    {
      title: '识别过程',
      items: [
        { value: '识别中', label: '识别中' },
        { value: '识别失败', label: '识别失败' },
        { value: '识别超时', label: '识别超时' },
        { value: '未识别到签字位', label: '未识别到签字位' },
      ],
    },
    {
      title: '版式',
      items: [
        { value: '版式-左右', label: '版式-左右' },
        { value: '版式-上下', label: '版式-上下' },
        { value: '版式-混合', label: '版式-混合' },
        { value: '版式-同格', label: '版式-同格' },
        { value: '版式-分格', label: '版式-分格' },
        { value: '版式-正文', label: '版式-正文' },
      ],
    },
    {
      title: '分隔方式',
      items: [
        { value: '分隔-/', label: '分隔-/' },
        { value: '分隔-空格', label: '分隔-空格' },
        { value: '分隔-空格子', label: '分隔-空格子' },
        { value: '分隔-单元格', label: '分隔-单元格' },
        { value: '分隔-换行', label: '分隔-换行' },
      ],
    },
  ];
  /** 筛选项与行上 tag/结论 的等价匹配（避免「缺姓名位」与「姓名位缺失」对不上） */
  var WORKBENCH_SLOT_FILTER_MATCH_KEYS = {
    缺姓名位: ['缺姓名位', '姓名位缺失'],
    缺日期位: ['缺日期位', '日期位缺失'],
    '缺姓名+日期位': ['缺姓名+日期位', '姓名位缺失', '日期位缺失'],
    姓名位缺失: ['姓名位缺失', '缺姓名位', '缺姓名+日期位'],
    日期位缺失: ['日期位缺失', '缺日期位', '缺姓名+日期位'],
  };
  var __signUploadInFlight = false;
  var __fileCacheHydratePromise = null;
  /** aiword 批量载入后由工作台统一识别，renderFileList 勿再单文件自动识别 */
  var __aiwordPendingBatchWorkbench = false;
  var __batchWorkbenchAdvancedOpen = false;
  var __batchWorkbenchEditFileId = null;
  var currentRoleMap = {};
  var signersList = [];
  var signersDbShare = false;
  /** 素材录入「当前签署人」唯一来源（快捷按钮与下拉共用，避免仅改 select 未生效） */
  var libActiveSignerId = '';
  var libSignerQuickPickEl = null;
  var libLocaleQuickPickEl = null;

  var signerPageIndex = 0;
  var signerPageSize = 3;
  var signedListPage = 1;
  var signedListPageSize = 10;
  var signedListQ = '';
  var strokeItemPage = 1;
  /** 并发 refreshFileList 时忽略过期响应，避免列表一直停在「正在加载…」 */
  var _fileListRefreshGen = 0;
  // 已存储签字图片：默认 3 条/页（可翻页）
  var strokeItemPageSize = 3;
  var strokeItemQ = '';
  var strokeItemCat = '';
  var roleLocaleMap = {};
  var roleLocaleManual = {};
  // 本文件角色映射：每个角色的素材下拉框独立筛选（按签署人姓名/ID）
  var roleItemFilterQ = {};
  // 本文件角色映射：把“画布签名/日期”入库到哪个签署人（每个角色、每个 kind 可不同）
  var roleSaveSigner = {};

  function _deepCloneJsonish(x) {
    try {
      return JSON.parse(JSON.stringify(x || {}));
    } catch (_) {
      return {};
    }
  }

  function saveCurrentFileUiToCache(fileId) {
    if (!fileId) return;
    fileUiCache[fileId] = fileUiCache[fileId] || {};
    fileUiCache[fileId].detectedOnce = !!(
      (fileUiCache[fileId] && fileUiCache[fileId].detectedOnce) ||
      (lastDetectData && lastDetectData.ok && String(lastDetectFileId) === String(fileId))
    );
    fileUiCache[fileId].lastDetectData =
      lastDetectData && String(lastDetectFileId) === String(fileId) ? lastDetectData : (fileUiCache[fileId].lastDetectData || null);
    fileUiCache[fileId].lastDetectError =
      String(lastDetectFileId) === String(fileId) ? (lastDetectError || '') : (fileUiCache[fileId].lastDetectError || '');
    fileUiCache[fileId].checkedRoleIds = selectedRoleIds();
    fileUiCache[fileId].roleItemFilterQ = _deepCloneJsonish(roleItemFilterQ);
    fileUiCache[fileId].roleSaveSigner = _deepCloneJsonish(roleSaveSigner);
    fileUiCache[fileId].currentRoleMap = _deepCloneJsonish(currentRoleMap);
  }

  function restoreFileUiFromCache(fileId) {
    var st = fileId ? fileUiCache[fileId] : null;
    if (!st) return false;
    // 检测结果
    lastDetectData = st.lastDetectData || null;
    lastDetectFileId = st.lastDetectData ? fileId : null;
    lastDetectError = st.lastDetectError || '';
    // 勾选角色（不强制要求 detect 成功）
    resetAllRoleChecks();
    (st.checkedRoleIds || []).forEach(function (rid) {
      if (rid) setRoleChecked(rid, true);
    });
    // 表格筛选/入库绑定等（体验：切回来不丢输入）
    roleItemFilterQ = _deepCloneJsonish(st.roleItemFilterQ);
    roleSaveSigner = _deepCloneJsonish(st.roleSaveSigner);
    currentRoleMap = _deepCloneJsonish(st.currentRoleMap);
    renderNeedSignTable();
    updateSubmitState();
    showNeedSignNoticeForSelectedFile();
    return true;
  }

  function cachePatchCurrentRoleMap(fileId, newMap) {
    if (!fileId) return;
    fileUiCache[fileId] = fileUiCache[fileId] || {};
    fileUiCache[fileId].currentRoleMap = _deepCloneJsonish(newMap || {});
  }

  /** 将「角色→素材」完整写入服务端（库映射签字读的是 MySQL，必须先落库再签） */
  function persistRoleMapToServer(fileId, mapOpt) {
    if (!fileId) return Promise.resolve(null);
    var m = mapOpt != null ? mapOpt : currentRoleMap;
    if (!m || typeof m !== 'object') m = {};
    return fetchJson(apiUrl('/api/sign/files/' + fileId + '/role-map'), {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ map: m }),
    }).then(function (r) {
      var jj = r.data || {};
      if (!jj.ok) throw new Error(jj.error || '保存角色映射失败');
      if (String(selectedFileId) === String(fileId)) {
        currentRoleMap = jj.map || m;
      }
      cachePatchCurrentRoleMap(fileId, jj.map || m);
      finalizeWorkbenchRowStatus(fileId);
      schedulePersistFileSessionCache(fileId);
      return jj.map || m;
    });
  }

  /** 库映射批量签：把各文件在内存里最后一次的映射刷到服务器，避免签的是旧配置。
   *  返回一个数组：每个文件的 { fid, ok, error }，调用方可据此告知用户哪些文件 role-map 落库失败。
   *  注意：本函数**不再 throw**，单文件 PUT 失败也继续下一文件，并把失败信息透传给调用方。 */
  function bulkPersistRoleMapsForBatch(fileIds) {
    var chain = Promise.resolve();
    var report = [];
    (fileIds || []).forEach(function (fid) {
      chain = chain.then(function () {
        if (!fid) {
          report.push({ fid: fid, ok: false, error: '无效 file_id' });
          return null;
        }
        var m =
          String(fid) === String(selectedFileId)
            ? _deepCloneJsonish(currentRoleMap || {})
            : _deepCloneJsonish((fileUiCache[fid] || {}).currentRoleMap || {});
        if (!m || !Object.keys(m).length) {
          // cache 为空时不主动 PUT 空对象（避免覆盖服务端已有 map）；
          // 但记录一条 warning，让用户知道这个文件依赖服务端旧 role-map。
          report.push({ fid: fid, ok: true, warning: 'cache 为空，使用服务端已有映射' });
          return null;
        }
        return persistRoleMapToServer(fid, m)
          .then(function () { report.push({ fid: fid, ok: true }); })
          .catch(function (e) {
            report.push({ fid: fid, ok: false, error: (e && e.message) || String(e || '') });
          });
      });
    });
    return chain.then(function () { return report; });
  }

  function cacheSetNeedSignNotice(fileId, msg, isErr) {
    if (!fileId) return;
    fileUiCache[fileId] = fileUiCache[fileId] || {};
    fileUiCache[fileId].needSignNotice = msg || '';
    fileUiCache[fileId].needSignNoticeErr = !!isErr;
  }

  function showNeedSignNoticeForSelectedFile() {
    try {
      if (!selectedFileId) return;
      var st = fileUiCache[selectedFileId] || {};
      var msg = st.needSignNotice || '';
      if (!msg) {
        setNeedSignActionFeedback('');
        return;
      }
      setNeedSignActionFeedback(msg, !!st.needSignNoticeErr);
    } catch (_) {}
  }

  function cacheDetectResultForFile(fileId, detectJson, errMsg) {
    if (!fileId) return;
    fileUiCache[fileId] = fileUiCache[fileId] || {};
    fileUiCache[fileId].detectedOnce = true;
    if (detectJson && detectJson.ok) {
      var bad = validateDetectResponseForFile(fileId, detectJson);
      if (bad) {
        fileUiCache[fileId].lastDetectData = {
          ok: false,
          error: bad,
          file_id: fileId,
        };
        fileUiCache[fileId].lastDetectError = bad;
      } else {
        fileUiCache[fileId].lastDetectData = detectJson;
        fileUiCache[fileId].lastDetectError = '';
        syncWorkbenchRowFromDetect(fileId, detectJson);
      }
    } else if (errMsg) {
      fileUiCache[fileId].lastDetectError = String(errMsg);
      fileUiCache[fileId].lastDetectData = {
        ok: false,
        error: String(errMsg),
        file_id: fileId,
      };
    } else if (detectJson && !detectJson.ok) {
      fileUiCache[fileId].lastDetectData = detectJson;
      fileUiCache[fileId].lastDetectError = String(detectJson.error || '识别失败');
    }
    if (String(selectedFileId) === String(fileId)) {
      fileUiCache[fileId].checkedRoleIds = selectedRoleIds();
    }
    schedulePersistFileSessionCache(fileId);
  }

  function cacheMarkDetected(fileId) {
    if (!fileId) return;
    if (lastDetectData && lastDetectData.ok && String(lastDetectFileId) === String(fileId)) {
      cacheDetectResultForFile(fileId, lastDetectData, '');
    } else if (lastDetectError && String(lastDetectFileId) === String(fileId)) {
      cacheDetectResultForFile(fileId, null, lastDetectError);
    }
    fileUiCache[fileId] = fileUiCache[fileId] || {};
    fileUiCache[fileId].detectedOnce = true;
    if (String(selectedFileId) === String(fileId)) {
      fileUiCache[fileId].checkedRoleIds = selectedRoleIds();
    }
  }

  function isCompositeDateMode(dm) {
    var d = String(dm || '').toLowerCase();
    return (
      d === 'composite_zh_ymd' || d === 'composite_en_space' || d === 'composite_en'
    );
  }

  function compositeModeToLayout(dm) {
    var d = String(dm || '').toLowerCase();
    if (d === 'composite_zh_ymd') return 'zh_ymd';
    if (d === 'composite_en_space') return 'en_space';
    // 兼容旧值 composite_en：前端不再提供「15.April.2026」，统一映射为英文空格版
    if (d === 'composite_en') return 'en_space';
    return 'en_space';
  }

  /** 与后端 date_piece_compose.kinds_* 对齐：给定日期与版式所需元件 kind 列表 */
  function compositeRequiredPieceKinds(dateMode, dateIso) {
    if (!isCompositeDateMode(dateMode)) return [];
    var iso = _parseAiwordDocDateIso(dateIso);
    if (!iso) return [];
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
    if (!m) return [];
    var y = m[1];
    var mo = m[2];
    var d = m[3];
    var moN = parseInt(mo, 10);
    var lay = compositeModeToLayout(dateMode);
    var kinds = [];
    var ch;
    function pushDigits(s) {
      for (var i = 0; i < s.length; i++) {
        ch = s.charAt(i);
        if (ch >= '0' && ch <= '9') kinds.push('pd' + ch);
      }
    }
    if (lay === 'zh_ymd') {
      pushDigits(y);
      kinds.push('pdot');
      pushDigits(mo);
      kinds.push('pdot');
      pushDigits(d);
    } else if (lay === 'en_space') {
      pushDigits(String(parseInt(d, 10)));
      if (moN >= 1 && moN <= 12) {
        kinds.push('pma' + ('0' + moN).slice(-2));
      }
      pushDigits(y);
    } else {
      pushDigits(String(parseInt(d, 10)));
      kinds.push('pdot');
      if (moN >= 1 && moN <= 12) {
        kinds.push('pm' + ('0' + moN).slice(-2));
      }
      kinds.push('pdot');
      pushDigits(y);
    }
    return kinds;
  }

  function signerHasCompositePieceKind(signer, kind) {
    if (!signer || !kind) return false;
    var pie = signer.date_piece_en;
    if (pie && typeof pie === 'object' && pie[kind]) return true;
    return false;
  }

  /** 返回缺失的元件 kind；空数组表示库内元件齐全（与 date_piece_en 一致） */
  function signerMissingCompositePieceKinds(signerId, dateMode, dateIso) {
    var signer = _findSignerById(signerId);
    if (!signer) return ['签署人未在列表中'];
    var required = compositeRequiredPieceKinds(dateMode, dateIso);
    var missing = [];
    for (var i = 0; i < required.length; i++) {
      var k = required[i];
      if (!signerHasCompositePieceKind(signer, k)) {
        missing.push(k);
      }
    }
    return missing;
  }

  function roleMapEntryNonEmpty(p) {
    if (!p || typeof p !== 'object') return false;
    if (isCompositeDateMode(p.date_mode)) {
      // 拼接日期是“配置型”选项：允许先选 date_mode/date_iso，后续再补选签名素材。
      // 否则 putRoleMap 会把该角色从 map 删除，表格重绘（如点“批量映射”）时看起来像“被清空”。
      return !!(p.date_iso || p.sig);
    }
    return !!(p.sig || p.date);
  }

  /** 库映射自动模式：映射表中已绑定且结构满足生成条件的角色（顺序与 ROLES 一致） */
  function libraryBoundSignableRoleIds(map) {
    var m = map && typeof map === 'object' ? map : {};
    var out = [];
    ROLES.forEach(function (r) {
      var p = m[r.id];
      if (p && typeof p === 'object' && roleMapEntryNonEmpty(p)) {
        out.push(r.id);
      }
    });
    return out;
  }

  function signSourceValue() {
    return (signSourceMode && signSourceMode.value ? signSourceMode.value : 'canvas').trim().toLowerCase();
  }

  function libraryRolesRestrictedToChecks() {
    return !!(libraryRolesUseChecksCb && libraryRolesUseChecksCb.checked);
  }

  /** 单文件生成：画布始终用勾选角色；库映射默认用表绑定，高级选项用勾选 */
  function effectiveRolesForSingleSign() {
    if (signSourceValue() === 'canvas') return selectedRoleIds();
    if (libraryRolesRestrictedToChecks()) return selectedRoleIds();
    return libraryBoundSignableRoleIds(currentRoleMap);
  }

  function syncLibraryRolesModeRow() {
    if (!libraryRolesModeRow || !signSourceMode) return;
    libraryRolesModeRow.style.display = signSourceValue() === 'library' ? 'block' : 'none';
  }

  function todayIsoLocal() {
    var d = new Date();
    var y = d.getFullYear();
    var m = d.getMonth() + 1;
    var day = d.getDate();
    return y + '-' + (m < 10 ? '0' : '') + m + '-' + (day < 10 ? '0' : '') + day;
  }

  /** 批量工作台：拼接日期输入框初值；无 aiword 文档体现日期时不填今天 */
  function _workbenchCompositeDateIso(row, pair) {
    var p = pair && typeof pair === 'object' ? pair : {};
    if (p.date_iso) {
      var a = _formatDateInputValue(p.date_iso);
      if (a) return a;
    }
    if (row && row.doc_date) {
      var b = _formatDateInputValue(row.doc_date);
      if (b) return b;
    }
    return '';
  }

  /** 单文件页：来自 aiword 且无文档体现日期时留空；其它场景仍可用今天作兜底 */
  function _singleFileCompositeDateIso(pair) {
    var p = pair && typeof pair === 'object' ? pair : {};
    if (p.date_iso) {
      var a = _formatDateInputValue(p.date_iso);
      if (a) return a;
    }
    if (_isFromAiwordHandoff()) {
      return _formatDateInputValue((__aiwordHandoffCtx || {}).doc_date || '') || '';
    }
    return todayIsoLocal();
  }

  /** 将接口返回的 author：… 转为中文角色名 */
  function _humanizeRoleResultField(txt) {
    var s = String(txt || '').trim();
    if (!s) return '';
    return s
      .split(/[；;]/)
      .map(function (part) {
        part = String(part || '').trim();
        if (!part) return '';
        var m = /^([a-z][a-z0-9_]*)\s*[：:]\s*(.*)$/i.exec(part);
        if (!m) return part;
        var tail = m[2] || '';
        tail = tail
          .replace(/\bauthor\b/gi, roleLabel('author'))
          .replace(/\breviewer\b/gi, roleLabel('reviewer'))
          .replace(/\bapprover\b/gi, roleLabel('approver'))
          .replace(/\bexecutor\b/gi, roleLabel('executor'));
        return roleLabel(m[1]) + '：' + tail;
      })
      .filter(Boolean)
      .join('；');
  }

  function _fileDisplayNameById(fileId) {
    var rec = savedFiles.find(function (x) {
      return x && String(x.id) === String(fileId);
    });
    return (rec && rec.name) || String(fileId);
  }

  /** 生成前检查：缺素材 / 拼接笔迹风险；至少一项可签则允许继续 */
  function buildSignPreflightForFiles(fileIds) {
    var blockers = [];
    var partial = [];
    var compositeWarn = [];
    var canProceed = false;
    (fileIds || []).forEach(function (fid) {
      var mat = assessFileMaterialStatus(fid);
      var nm = _fileDisplayNameById(fid);
      var hasLib = (mat.matched || 0) > 0;
      var hasCanvas = (mat.canvas || 0) > 0;
      if (!hasLib && !hasCanvas) {
        blockers.push(nm + '（无任何角色已绑定可用素材）');
        return;
      }
      canProceed = true;
      var map = getFileRoleMapForWorkbench(fid);
      (mat.missingRoleIds || []).forEach(function (rid) {
        partial.push(nm + ' · ' + roleLabel(rid) + '：缺少签名/日期素材');
      });
      (mat.sigOnlyRoleIds || []).forEach(function (rid) {
        var p = map[rid] || {};
        if (isCompositeDateMode(p.date_mode) && !p.date_iso) {
          partial.push(nm + ' · ' + roleLabel(rid) + '：已选签名，未填文档体现日期');
        } else {
          partial.push(nm + ' · ' + roleLabel(rid) + '：仅签名素材，日期将跳过');
        }
      });
      (mat.dateOnlyRoleIds || []).forEach(function (rid) {
        partial.push(nm + ' · ' + roleLabel(rid) + '：仅日期素材，签名将跳过');
      });
      (mat.roleDetails || []).forEach(function (d) {
        if (!d || d.state !== 'full') return;
        var p = map[d.id] || {};
        if (!isCompositeDateMode(p.date_mode) || !p.date_iso || !p.sig) return;
        var sid = _strokeItemSignerId('sig', p.sig);
        var signerNm = sid && _findSignerById(sid) ? _findSignerById(sid).name : '';
        var missKinds = sid ? signerMissingCompositePieceKinds(sid, p.date_mode, p.date_iso) : ['未绑定签名素材'];
        if (!missKinds.length) return;
        compositeWarn.push(
          nm +
            ' · ' +
            roleLabel(d.id) +
            '：拼接日期' +
            (signerNm ? '（' + signerNm + '）' : '') +
            '，缺少数字/月份元件：' +
            missKinds.slice(0, 8).join('、') +
            (missKinds.length > 8 ? '…' : '')
        );
      });
    });
    return {
      blockers: blockers,
      partial: partial,
      compositeWarn: compositeWarn,
      canProceed: canProceed,
    };
  }

  function confirmPartialSignProceed(preflight) {
    preflight = preflight || { blockers: [], partial: [], compositeWarn: [] };
    var lines = [];
    if (preflight.partial && preflight.partial.length) {
      lines.push('以下字段素材不完整，生成时将跳过或仅处理已匹配部分：');
      preflight.partial.forEach(function (x) {
        lines.push('  · ' + x);
      });
    }
    if (preflight.compositeWarn && preflight.compositeWarn.length) {
      lines.push('以下字段使用拼接日期，数字笔迹未录入时可能无法签日期（仅签签名或跳过）：');
      preflight.compositeWarn.forEach(function (x) {
        lines.push('  · ' + x);
      });
    }
    lines.push('');
    lines.push('是否仍继续生成已签名文档？');
    return window.confirm(lines.join('\n'));
  }

  function formatBatchSignItemResult(it) {
    if (!it) return '';
    var nm = it.name || it.file_id || '';
    if (!it.ok) {
      var failParts = ['❌ ' + nm + '：' + (it.error || '失败')];
      if (Array.isArray(it.missing_roles) && it.missing_roles.length) {
        failParts.push('  未落位角色：' + it.missing_roles.join('、'));
      }
      if (it.per_role_results && typeof it.per_role_results === 'object') {
        var missLines = [];
        Object.keys(it.per_role_results).forEach(function (rid) {
          var one = it.per_role_results[rid] || {};
          if (one && one.placed) return;
          var why = one && one.placed_by ? String(one.placed_by) : 'not_found';
          missLines.push(rid + '（' + why + '）');
        });
        if (missLines.length) {
          failParts.push('  角色明细：' + missLines.join('；'));
        }
      }
      return failParts.join('\n');
    }
    var chunks = [];
    if (it.applied_n > 0) {
      chunks.push(
        '已成功 ' +
          it.applied_n +
          ' 项' +
          (it.applied ? '：' + _humanizeRoleResultField(it.applied) : '')
      );
    } else {
      chunks.push('已成功 0 项');
    }
    if (it.skipped_n > 0) {
      chunks.push(
        '未签/跳过 ' +
          it.skipped_n +
          ' 项' +
          (it.skipped ? '：' + _humanizeRoleResultField(it.skipped) : '')
      );
    }
    if (Array.isArray(it.fallback_roles) && it.fallback_roles.length) {
      chunks.push('规则建议：以下角色本次依赖兜底关键词落位，建议补充版式规则：' + it.fallback_roles.join('、'));
    }
    return '✅ ' + nm + '\n  ' + chunks.join('\n  ');
  }

  function formatSingleSignApplySummary(sum) {
    sum = sum || {};
    var apN = sum.applied_n || 0;
    var skN = sum.skipped_n || 0;
    var lines = ['签字结果摘要：'];
    if (apN > 0) {
      lines.push(
        '已成功 ' + apN + ' 项' + (sum.applied ? '：' + _humanizeRoleResultField(sum.applied) : '')
      );
    } else {
      lines.push('未插入任何签名/日期（全部跳过）。');
    }
    if (skN > 0) {
      lines.push(
        '未签/跳过 ' + skN + ' 项' + (sum.skipped ? '：' + _humanizeRoleResultField(sum.skipped) : '')
      );
    }
    return lines.join('\n');
  }

  function roleLocaleLabel(loc) {
    return loc === 'en' ? '英文' : '中文';
  }

  function localeFromStrokeSetOption(sel) {
    try {
      var opt = sel && sel.options ? sel.options[sel.selectedIndex] : null;
      if (!opt || !opt.textContent) return '';
      return /英文/.test(opt.textContent) ? 'en' : (/中文/.test(opt.textContent) ? 'zh' : '');
    } catch (_) {
      return '';
    }
  }

  function showSignerErr(s) {
    if (!signerErrMsg) return;
    if (s) {
      clearFileRegionErr();
      clearNeedSignScopedFeedbacks();
    }
    signerErrMsg.style.display = s ? 'block' : 'none';
    signerErrMsg.textContent = s || '';
    if (s) {
      signerErrMsg.style.color =
        /失败|错误|无效|无法|缺少|请先|未能/.test(s) ? 'var(--error)' : 'var(--text-muted)';
      try {
        if (window.matchMedia && window.matchMedia('(pointer: coarse)').matches) {
          requestAnimationFrame(function () {
            signerErrMsg.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
          });
        }
      } catch (_) {}
    }
  }

  function clearFileRegionErr() {
    if (!errMsg) return;
    errMsg.style.display = 'none';
    errMsg.textContent = '';
  }

  function setLibStrokeFeedback(s, isErr) {
    if (!libStrokeFeedback) return;
    if (!s) {
      libStrokeFeedback.style.display = 'none';
      libStrokeFeedback.textContent = '';
      return;
    }
    libStrokeFeedback.style.display = 'block';
    libStrokeFeedback.textContent = s;
    libStrokeFeedback.className = 'btn-inline-feedback' + (isErr ? ' is-error' : ' is-ok');
    if (isErr) {
      try {
        requestAnimationFrame(function () {
          libStrokeFeedback.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
        });
      } catch (_) {}
    }
  }

  function onLibStrokeImgErr(msg) {
    setLibStrokeFeedback(msg, true);
  }

  function setPieceBatchProgress(msg) {
    if (!pieceBatchStatus) return;
    pieceBatchStatus.textContent = msg || '';
    pieceBatchStatus.className = 'btn-inline-feedback is-ok';
  }

  function setPieceBatchFeedback(s, isErr) {
    if (!pieceBatchStatus) return;
    if (!s) {
      pieceBatchStatus.textContent = '';
      pieceBatchStatus.className = 'btn-inline-feedback is-ok';
      return;
    }
    pieceBatchStatus.textContent = s;
    pieceBatchStatus.className =
      'btn-inline-feedback' + (isErr ? ' is-error' : ' is-ok');
    if (isErr) {
      try {
        requestAnimationFrame(function () {
          pieceBatchStatus.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
        });
      } catch (_) {}
    }
  }

  function setPieceHintFeedback(s, isErr) {
    if (!pieceHint) return;
    if (!s) {
      pieceHint.textContent = '';
      pieceHint.className = 'btn-inline-feedback is-ok';
      return;
    }
    pieceHint.textContent = s;
    pieceHint.className =
      'btn-inline-feedback' + (isErr ? ' is-error' : ' is-ok');
    if (isErr) {
      try {
        requestAnimationFrame(function () {
          pieceHint.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
        });
      } catch (_) {}
    }
  }

  function setStrokeItemsActionFeedback(s, isErr) {
    if (!strokeItemsActionMsg) return;
    if (!s) {
      strokeItemsActionMsg.style.display = 'none';
      strokeItemsActionMsg.textContent = '';
      return;
    }
    strokeItemsActionMsg.style.display = 'block';
    strokeItemsActionMsg.textContent = s;
    strokeItemsActionMsg.className =
      'btn-inline-feedback' + (isErr ? ' is-error' : ' is-ok');
  }

  function setFileListActionFeedback(s, isErr) {
    if (!fileListActionMsg) return;
    if (!s) {
      fileListActionMsg.style.display = 'none';
      fileListActionMsg.textContent = '';
      return;
    }
    fileListActionMsg.style.display = 'block';
    fileListActionMsg.textContent = s;
    fileListActionMsg.className =
      'btn-inline-feedback' + (isErr ? ' is-error' : ' is-ok');
  }

  function formatArchiveSummaryLine(sum) {
    var s = sum && typeof sum === 'object' ? sum : null;
    if (!s) return '';
    var nArc = Number(s.archives) || 0;
    if (!nArc) return '';
    var total = Number(s.total_members) || 0;
    var source = Number(s.source_members) || 0;
    var added = Number(s.added_signable) || 0;
    var skipped = Math.max(0, total - source);
    var extMap = (s.skipped_by_ext && typeof s.skipped_by_ext === 'object') ? s.skipped_by_ext : {};
    var tops = Object.keys(extMap)
      .sort(function (a, b) { return (Number(extMap[b]) || 0) - (Number(extMap[a]) || 0); })
      .slice(0, 4)
      .map(function (k) { return k + ':' + (Number(extMap[k]) || 0); })
      .join('，');
    var line =
      '压缩包统计：总文件 ' + total +
      '，可签字候选 ' + source +
      '，成功入列 ' + added +
      '，跳过 ' + skipped;
    if (tops) line += '（按扩展名：' + tops + '）';
    return line;
  }

  function batchDeleteSelectedFiles(feedbackTarget) {
    var ids = _checkedBatchFileIds();
    if (!ids.length) {
      if (feedbackTarget === 'workbench') setBatchWorkbenchMsg('请先勾选要删除的文件', true);
      else setFileListActionFeedback('请先勾选要删除的文件', true);
      return Promise.resolve();
    }
    beginPageProgress('正在删除 ' + ids.length + ' 个文件…', { done: 0, total: ids.length });
    return _batchDeleteSelectedFilesCore(feedbackTarget, ids).finally(function () {
      endPageProgress();
    });
  }

  function _batchDeleteSelectedFilesCore(feedbackTarget, ids) {
    var tip =
      '确定删除已勾选的 ' +
      ids.length +
      ' 个文件？\n删除后会同时清除这些文件的角色映射缓存。';
    if (!window.confirm(tip)) return Promise.resolve();
    var delSetReq = {};
    ids.forEach(function (x) {
      delSetReq[String(x)] = true;
    });
    // 先本地移除，减少大批量删除时的“等待无反馈”感；失败后会自动从服务端刷新纠偏。
    savedFiles = savedFiles.filter(function (rec) {
      return !(rec && rec.id && delSetReq[String(rec.id)]);
    });
    if (selectedFileId && delSetReq[String(selectedFileId)]) {
      selectedFileId = savedFiles.length ? String(savedFiles[0].id || '') : null;
    }
    Object.keys(delSetReq).forEach(function (x) {
      try {
        delete fileUiCache[String(x)];
        delete __batchWorkbenchRows[String(x)];
        delete __aiwordHandoffCtxByFileId[String(x)];
        delete _lastPersistedRoleMapStable[String(x)];
      } catch (_) {}
    });
    renderFileList();
    syncHiddenBatchPicks();
    renderBatchWorkbenchTable();
    if (feedbackTarget === 'workbench') {
      setBatchWorkbenchMsg('正在删除 ' + ids.length + ' 个文件…', false);
    } else {
      setFileListActionFeedback('正在删除 ' + ids.length + ' 个文件…', false);
    }
    var post = function () {
      return fetchJson(apiUrl('/api/sign/files/batch-delete'), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ file_ids: ids }),
        timeoutMs: 300000,
      }).then(function (r) {
        var j = (r && r.data) || {};
        if (!j.ok) {
          var em = j.error || '批量删除失败';
          if (feedbackTarget === 'workbench') setBatchWorkbenchMsg(em, true);
          else setFileListActionFeedback(em, true);
          throw new Error(em);
        }
        var deleted = Array.isArray(j.deleted_ids) ? j.deleted_ids.map(String) : [];
        var delSet = {};
        deleted.forEach(function (x) {
          delSet[String(x)] = true;
          try {
            delete fileUiCache[String(x)];
            delete __batchWorkbenchRows[String(x)];
            delete __aiwordHandoffCtxByFileId[String(x)];
            delete _lastPersistedRoleMapStable[String(x)];
          } catch (_) {}
        });
        savedFiles = normalizeSavedFileRecords(Array.isArray(j.files) ? j.files : []).filter(function (rec) {
          return !(rec && rec.id && delSet[String(rec.id)]);
        });
        if (selectedFileId && delSet[String(selectedFileId)]) {
          selectedFileId = savedFiles.length ? String(savedFiles[0].id || '') : null;
        }
        renderFileList();
        syncHiddenBatchPicks();
        renderBatchWorkbenchTable();
        var msg = '已删除 ' + deleted.length + ' 个文件';
        var miss = Array.isArray(j.missing_ids) ? j.missing_ids.length : 0;
        if (miss) msg += '（' + miss + ' 个不存在）';
        if (feedbackTarget === 'workbench') setBatchWorkbenchMsg(msg, false);
        else setFileListActionFeedback(msg, false);
      });
    };
    return post().catch(function (e) {
      var em = (e && e.message) || String(e || '');
      return refreshFileList({ softFail: false, silent: true, skipPageProgress: true })
        .catch(function () {})
        .then(function () {
          var msg = '批量删除请求失败：' + em + '。已自动刷新列表。';
          if (feedbackTarget === 'workbench') setBatchWorkbenchMsg(msg, true);
          else setFileListActionFeedback(msg, true);
        });
    });
  }

  function setSignedListActionFeedback(s, isErr) {
    if (!signedListActionMsg) return;
    if (!s) {
      signedListActionMsg.style.display = 'none';
      signedListActionMsg.textContent = '';
      return;
    }
    signedListActionMsg.style.display = 'block';
    signedListActionMsg.textContent = s;
    signedListActionMsg.className =
      'btn-inline-feedback' + (isErr ? ' is-error' : ' is-ok');
  }

  function parseSignerNamesInput(raw) {
    return raw
      .split(/[,，;；\r\n]+/)
      .map(function (x) {
        return x.trim();
      })
      .filter(function (x) {
        return x.length > 0;
      });
  }

  /**
   * 与 /api/sign/signers 一致。素材 = 已保存的手写 PNG（签名图 / 日期图，可成套或分条）。
   */
  function signerAssetsBrief(s) {
    s = s || {};
    var sets = (s.stroke_sets || []).length;
    var sigs = (s.sig_items || []).length;
    var dates = (s.date_items || []).length;
    var parts = [];
    if (sets) parts.push(sets + '套成对（签名+日期）');
    if (sigs) parts.push('签名手写图 ' + sigs + ' 张');
    if (dates) parts.push('日期手写图 ' + dates + ' 张');
    var hs = !!s.has_sig;
    var hd = !!s.has_date;
    var statusShort;
    var statusLine;
    if (hs && hd) {
      statusShort = '签名与日期均已备';
      statusLine = '生成已签文档：签名与日期素材均已备齐。';
    } else if (hs && !hd) {
      statusShort = '还缺日期';
      statusLine = '生成已签文档：还缺「日期」手写图（请在「日期」画布书写并保存）。';
    } else if (!hs && hd) {
      statusShort = '还缺签名';
      statusLine = '生成已签文档：还缺「签名」手写图（请在「签名」画布书写并保存）。';
    } else {
      statusShort = '签名与日期均未入库';
      statusLine = '生成已签文档：签名与日期手写图均未入库。';
    }
    return {
      hint: parts.length ? parts.join(' · ') : '尚无已保存的手写图',
      ready: !!(hs && hd),
      statusShort: statusShort,
      statusLine: statusLine,
    };
  }

  /** 移动端常见：select.value 已与 selectedIndex 指向的 option 不一致 */
  function _syncLibSignerSelectIndex(sid) {
    if (!libSignerSelect || !sid) return;
    for (var i = 0; i < libSignerSelect.options.length; i++) {
      var op = libSignerSelect.options[i];
      if (op && op.value === sid) {
        libSignerSelect.selectedIndex = i;
        return;
      }
    }
  }

  function _libSignerFeedbackLabel(sid) {
    var s = _findSignerById(sid);
    var nm = s ? String(s.name || s.id || sid) : String(sid || '');
    var id = String(sid || '');
    if (!id) return nm;
    var shortId = id.length > 10 ? id.slice(0, 8) + '…' : id;
    return nm + '（' + shortId + '）';
  }

  function getLibActiveSignerId() {
    var active = String(libActiveSignerId || '').trim();
    if (active) {
      if (!signersList || !signersList.length) return active;
      if (_findSignerById(active)) return active;
      return '';
    }
    var fromSel =
      libSignerSelect && libSignerSelect.value
        ? String(libSignerSelect.value).trim()
        : '';
    if (fromSel && _findSignerById(fromSel)) {
      libActiveSignerId = fromSel;
      _syncLibSignerSelectIndex(fromSel);
      return fromSel;
    }
    return '';
  }

  function setLibActiveSignerId(sid) {
    sid = String(sid || '').trim();
    if (!sid) {
      libActiveSignerId = '';
      if (libSignerSelect) libSignerSelect.value = '';
      syncLibStrokeSetSelect();
      syncCurrentSignerBanners();
      renderLibSignerQuickPick();
      return false;
    }
    if (!_findSignerById(sid)) return false;
    libActiveSignerId = sid;
    function _selectHasSigner() {
      if (!libSignerSelect) return false;
      return Array.prototype.some.call(libSignerSelect.options, function (o) {
        return o && o.value === sid;
      });
    }
    if (!_selectHasSigner() && libSignerFilter && String(libSignerFilter.value || '').trim()) {
      libSignerFilter.value = '';
      syncLibSignerSelect();
    }
    if (libSignerSelect && !_selectHasSigner()) {
      syncLibSignerSelect();
    }
    if (libSignerSelect) {
      libSignerSelect.value = sid;
      if (libSignerSelect.value !== sid && libSignerFilter) {
        libSignerFilter.value = '';
        syncLibSignerSelect();
        libSignerSelect.value = sid;
      }
      _syncLibSignerSelectIndex(sid);
    }
    syncLibStrokeSetSelect();
    syncCurrentSignerBanners();
    renderLibSignerQuickPick();
    return true;
  }

  function _libSignerSelectVisibleIds() {
    var out = [];
    if (!libSignerSelect) return out;
    for (var i = 0; i < libSignerSelect.options.length; i++) {
      var v = libSignerSelect.options[i] && libSignerSelect.options[i].value;
      if (v) out.push(String(v));
    }
    return out;
  }

  /** 筛选重建下拉后：仅保留仍可见的选中项；若筛后只剩一人则自动选中 */
  function _resolveLibSignerSelectValue(preferredId) {
    var visible = _libSignerSelectVisibleIds();
    if (preferredId && visible.indexOf(String(preferredId)) >= 0) {
      return String(preferredId);
    }
    if (visible.length === 1) {
      return visible[0];
    }
    return '';
  }

  function syncLibSignerSelect() {
    var prev = getLibActiveSignerId() || (libSignerSelect && libSignerSelect.value) || '';
    libSignerSelect.innerHTML = '';
    var o0 = document.createElement('option');
    o0.value = '';
    o0.textContent = signersList.length ? '请选择签署人（维护其签名/日期手写图）' : '请先在上文添加签署人';
    libSignerSelect.appendChild(o0);
    var q = (libSignerFilter && libSignerFilter.value) ? String(libSignerFilter.value).trim().toLowerCase() : '';
    signersList.forEach(function (s) {
      if (q) {
        var nm = (s && s.name ? String(s.name) : '').toLowerCase();
        var sid = (s && s.id ? String(s.id) : '').toLowerCase();
        if (nm.indexOf(q) < 0 && sid.indexOf(q) < 0) {
          return;
        }
      }
      var o = document.createElement('option');
      o.value = s.id;
      var brief = signerAssetsBrief(s);
      o.textContent =
        (s.name || s.id) + ' — ' + brief.hint + ' — ' + brief.statusShort;
      libSignerSelect.appendChild(o);
    });
    var resolved = _resolveLibSignerSelectValue(prev);
    libSignerSelect.value = resolved;
    libActiveSignerId = resolved || '';
    if (resolved) _syncLibSignerSelectIndex(resolved);
    renderLibSignerQuickPick();
    syncLibStrokeSetSelect();
    syncCurrentSignerBanners();
  }

  function ensureLibSignerQuickPickEl() {
    if (!libSignerSelect || !libSignerSelect.parentNode) return null;
    if (!libSignerQuickPickEl) {
      var host = document.createElement('div');
      host.id = 'libSignerQuickPick';
      host.style.marginTop = '8px';
      host.style.display = 'none';
      libSignerSelect.parentNode.insertBefore(host, libSignerSelect.nextSibling);
      libSignerQuickPickEl = host;
    }
    return libSignerQuickPickEl;
  }

  function renderLibSignerQuickPick() {
    var host = ensureLibSignerQuickPickEl();
    if (!host) return;
    var q = (libSignerFilter && libSignerFilter.value) ? String(libSignerFilter.value).trim().toLowerCase() : '';
    var sid = getLibActiveSignerId();
    var rows = [];
    for (var i = 0; i < signersList.length; i++) {
      var s = signersList[i] || {};
      var nm = (s && s.name ? String(s.name) : '');
      var id = (s && s.id ? String(s.id) : '');
      if (q) {
        var nmLow = nm.toLowerCase();
        var idLow = id.toLowerCase();
        if (nmLow.indexOf(q) < 0 && idLow.indexOf(q) < 0) continue;
      }
      rows.push({ id: id, name: nm || id });
    }
    if (!rows.length) {
      host.style.display = 'none';
      host.innerHTML = '';
      return;
    }
    host.style.display = 'block';
    var html = '<div class="hint" style="margin:6px 0;">手机端点下拉无响应时，可直接点这里选择签署人：</div><div style="display:flex;gap:6px;flex-wrap:wrap;">';
    for (var j = 0; j < rows.length; j++) {
      var r = rows[j];
      var active = sid && sid === r.id;
      html +=
        '<button type="button" class="btn btn-secondary lib-signer-chip" data-sid="' + r.id +
        '" style="padding:6px 10px;' + (active ? 'border-color:#1a73e8;color:#1a73e8;background:#eef4ff;' : '') + '">' +
        r.name + '</button>';
    }
    html += '</div>';
    host.innerHTML = html;
    var chips = host.querySelectorAll('.lib-signer-chip');
    for (var k = 0; k < chips.length; k++) {
      chips[k].addEventListener('click', function () {
        var picked = this.getAttribute('data-sid') || '';
        if (!picked) return;
        if (!setLibActiveSignerId(picked)) {
          showSignerErr('无法选中该签署人，请刷新列表后重试');
        } else {
          showSignerErr('');
        }
      });
    }
  }

  function ensureLibLocaleQuickPickEl() {
    if (!libLocaleSelect || !libLocaleSelect.parentNode) return null;
    if (!libLocaleQuickPickEl) {
      var host = document.createElement('div');
      host.id = 'libLocaleQuickPick';
      host.style.marginTop = '8px';
      host.style.display = 'none';
      libLocaleSelect.parentNode.insertBefore(host, libLocaleSelect.nextSibling);
      libLocaleQuickPickEl = host;
    }
    return libLocaleQuickPickEl;
  }

  function renderLibLocaleQuickPick() {
    var host = ensureLibLocaleQuickPickEl();
    if (!host || !libLocaleSelect) return;
    var cur = String(libLocaleSelect.value || 'zh');
    var opts = Array.prototype.slice.call(libLocaleSelect.options || []).filter(function (op) {
      return !!(op && op.value);
    });
    if (!opts.length) {
      host.style.display = 'none';
      host.innerHTML = '';
      return;
    }
    host.style.display = 'block';
    var html = '<div class="hint" style="margin:6px 0;">手机端点下拉无响应时，可直接点这里切换版本：</div><div style="display:flex;gap:6px;flex-wrap:wrap;">';
    for (var i = 0; i < opts.length; i++) {
      var op = opts[i];
      var v = String(op.value || '');
      var txt = String(op.textContent || v || '');
      var active = cur === v;
      html +=
        '<button type="button" class="btn btn-secondary lib-locale-chip" data-loc="' + v +
        '" style="padding:6px 10px;' + (active ? 'border-color:#1a73e8;color:#1a73e8;background:#eef4ff;' : '') + '">' +
        txt + '</button>';
    }
    html += '</div>';
    host.innerHTML = html;
    var chips = host.querySelectorAll('.lib-locale-chip');
    for (var j = 0; j < chips.length; j++) {
      chips[j].addEventListener('click', function () {
        var picked = this.getAttribute('data-loc') || '';
        if (!picked || !libLocaleSelect) return;
        libLocaleSelect.value = picked;
        renderLibLocaleQuickPick();
      });
    }
  }

  function syncCurrentSignerBanners() {
    var sid = getLibActiveSignerId();
    if (!sid) {
      currentSignerName.textContent = '—';
      pieceCurrentSignerName.textContent = '—';
      return;
    }
    var s = signersList.find(function (x) {
      return x.id === sid;
    });
    var nm = s ? (s.name || s.id) : sid;
    // 显示更清楚：姓名（id）
    var label = s && s.name ? (String(s.name) + '（' + sid + '）') : sid;
    currentSignerName.textContent = label;
    pieceCurrentSignerName.textContent = label;
  }

  function syncLibStrokeSetSelect() {
    var prev = libStrokeSetSelect.value;
    libStrokeSetSelect.innerHTML = '';
    var o0 = document.createElement('option');
    o0.value = '';
    o0.textContent = '不指定成对套（载入该人最近一套签名+日期）';
    libStrokeSetSelect.appendChild(o0);
    var sid = getLibActiveSignerId();
    var s = signersList.find(function (x) {
      return x.id === sid;
    });
    (s && s.stroke_sets ? s.stroke_sets : []).forEach(function (st) {
      var o = document.createElement('option');
      o.value = st.id;
      o.textContent =
        (st.label || '') + (st.updated_at ? ' · ' + st.updated_at : '');
      libStrokeSetSelect.appendChild(o);
    });
    if (prev) {
      var has = Array.prototype.some.call(libStrokeSetSelect.options, function (op) {
        return op.value === prev;
      });
      if (has) libStrokeSetSelect.value = prev;
    }
  }

  function signerIdFromStrokeSetSelect(sel) {
    var opt = sel.options[sel.selectedIndex];
    if (!opt || !opt.getAttribute) return '';
    return opt.getAttribute('data-signer-id') || '';
  }

  function fillRoleItemSelect(sel, kind, currentId, filterQ) {
    sel.innerHTML = '';
    var o0 = document.createElement('option');
    o0.value = '';
    o0.textContent = kind === 'date' ? '请选择日期素材' : '请选择签名素材';
    sel.appendChild(o0);
    var q = filterQ ? String(filterQ).trim().toLowerCase() : '';
    signersList.forEach(function (s) {
      if (q) {
        var nm = (s && s.name ? String(s.name) : '').toLowerCase();
        var sid = (s && s.id ? String(s.id) : '').toLowerCase();
        if (nm.indexOf(q) < 0 && sid.indexOf(q) < 0) {
          // 若只想按签署人筛选：直接跳过该人全部素材
          return;
        }
      }
      var arr = kind === 'date' ? (s.date_items || []) : (s.sig_items || []);
      (arr || []).forEach(function (st) {
        var o = document.createElement('option');
        o.value = st.id;
        o.setAttribute('data-signer-id', s.id);
        var tail = st.updated_at ? ' · ' + st.updated_at : '';
        var loc = st.locale === 'en' ? '英文' : '中文';
        o.textContent =
          (s.name || s.id) +
          ' · ' +
          loc +
          ' · ' +
          (kind === 'date' ? '日期' : '签名') +
          ' · ' +
          (st.label || '') +
          tail;
        sel.appendChild(o);
      });
    });
    if (currentId) {
      var ok = Array.prototype.some.call(sel.options, function (op) {
        return op.value === currentId;
      });
      if (ok) sel.value = currentId;
    }
  }

  function _ensureRoleFilterState(rid) {
    if (!roleItemFilterQ[rid] || typeof roleItemFilterQ[rid] !== 'object') {
      roleItemFilterQ[rid] = { sig: '', date: '' };
    } else {
      if (typeof roleItemFilterQ[rid].sig !== 'string') roleItemFilterQ[rid].sig = '';
      if (typeof roleItemFilterQ[rid].date !== 'string') roleItemFilterQ[rid].date = '';
    }
    return roleItemFilterQ[rid];
  }

  function _ensureRoleSaveSignerState(rid) {
    if (!roleSaveSigner[rid] || typeof roleSaveSigner[rid] !== 'object') {
      roleSaveSigner[rid] = { sig: { id: '', q: '' }, date: { id: '', q: '' } };
    }
    if (!roleSaveSigner[rid].sig || typeof roleSaveSigner[rid].sig !== 'object') {
      roleSaveSigner[rid].sig = { id: '', q: '' };
    }
    if (!roleSaveSigner[rid].date || typeof roleSaveSigner[rid].date !== 'object') {
      roleSaveSigner[rid].date = { id: '', q: '' };
    }
    if (typeof roleSaveSigner[rid].sig.id !== 'string') roleSaveSigner[rid].sig.id = '';
    if (typeof roleSaveSigner[rid].sig.q !== 'string') roleSaveSigner[rid].sig.q = '';
    if (typeof roleSaveSigner[rid].date.id !== 'string') roleSaveSigner[rid].date.id = '';
    if (typeof roleSaveSigner[rid].date.q !== 'string') roleSaveSigner[rid].date.q = '';
    return roleSaveSigner[rid];
  }

  function _roleFinalLocale(rid) {
    var v = (roleLocaleMap && roleLocaleMap[rid]) ? roleLocaleMap[rid] : 'auto';
    if (v === 'zh' || v === 'en') return v;
    return (libLocaleSelect && libLocaleSelect.value) ? libLocaleSelect.value : 'zh';
  }

  function _signerHasKindInLocale(signer, kind, loc) {
    if (!signer) return false;
    kind = (kind || '').toLowerCase();
    loc = (loc === 'en') ? 'en' : 'zh';
    if (signer.brief) {
      if (kind === 'sig') return !!signer.has_sig;
      if (kind === 'date') return !!signer.has_date;
      return false;
    }
    try {
      var sets = signer.stroke_sets || [];
      if (Array.isArray(sets) && sets.length) {
        for (var i = 0; i < sets.length; i++) {
          var st = sets[i] || {};
          if ((st.locale || 'zh') !== loc) continue;
          if (kind === 'sig' && st.has_sig_blob) return true;
          if (kind === 'date' && st.has_date_blob) return true;
        }
      }
    } catch (_) {}
    var arr = kind === 'date' ? (signer.date_items || []) : (signer.sig_items || []);
    if (!Array.isArray(arr)) return false;
    for (var j = 0; j < arr.length; j++) {
      var it = arr[j] || {};
      if ((it.locale || 'zh') === loc) return true;
    }
    return false;
  }

  function fillSignerSelect(sel, currentId, filterQ) {
    sel.innerHTML = '';
    var o0 = document.createElement('option');
    o0.value = '';
    o0.textContent = signersList.length ? '选择要保存到的签署人…' : '请先添加签署人';
    sel.appendChild(o0);
    var q = filterQ ? String(filterQ).trim().toLowerCase() : '';
    signersList.forEach(function (s) {
      if (q) {
        var nm = (s && s.name ? String(s.name) : '').toLowerCase();
        var sid = (s && s.id ? String(s.id) : '').toLowerCase();
        if (nm.indexOf(q) < 0 && sid.indexOf(q) < 0) return;
      }
      var o = document.createElement('option');
      o.value = s.id;
      var hsZh = _signerHasKindInLocale(s, 'sig', 'zh');
      var hdZh = _signerHasKindInLocale(s, 'date', 'zh');
      var hsEn = _signerHasKindInLocale(s, 'sig', 'en');
      var hdEn = _signerHasKindInLocale(s, 'date', 'en');
      o.textContent =
        (s.name || s.id) +
        '（中文 签名:' + (hsZh ? '有' : '无') + ' 日期:' + (hdZh ? '有' : '无') +
        '｜英文 签名:' + (hsEn ? '有' : '无') + ' 日期:' + (hdEn ? '有' : '无') +
        '）';
      sel.appendChild(o);
    });
    if (currentId) {
      var ok = Array.prototype.some.call(sel.options, function (op) {
        return op.value === currentId;
      });
      if (ok) sel.value = currentId;
    }
  }

  function refreshRoleSaveSignerControls() {
    if (!IS_FILE_SIGN_PAGE) return;
    // 角色画布旁的“保存到签署人”控件在 buildUI 时创建；需要在 signersList 更新后回填选项
    try {
      document.querySelectorAll('.role-save-signer').forEach(function (wrap) {
        var rid = wrap.getAttribute('data-rid') || '';
        var kind = wrap.getAttribute('data-kind') || '';
        if (!rid || !kind) return;
        kind = kind === 'date' ? 'date' : 'sig';
        var stAll = _ensureRoleSaveSignerState(rid);
        var st = stAll[kind] || { id: '', q: '' };
        var filter = wrap.querySelector('input.role-save-signer-filter');
        var sel = wrap.querySelector('select.role-save-signer-select');
        if (!sel) return;
        var q = filter ? (filter.value || st.q || '') : (st.q || '');
        if (filter) filter.value = q;
        fillSignerSelect(sel, st.id || sel.value || '', q);
      });
    } catch (_) {}
  }

  function loadLibStrokesFromServer() {
    var sid = getLibActiveSignerId();
    if (!sid) {
      setLibStrokeFeedback('请先在「当前签署人」中选择一位', true);
      return Promise.reject(new Error('no_signer'));
    }
    setLibStrokeFeedback('', false);
    var ts = '?t=' + Date.now();
    var locRaw = (libLocaleSelect && libLocaleSelect.value) ? String(libLocaleSelect.value).trim() : 'zh';
    var locParam = '&locale=' + encodeURIComponent(locRaw === 'en' ? 'en' : 'zh');
    var setId = libStrokeSetSelect.value;
    var urlSig =
      setId && setId.length === 32
        ? apiUrl('/api/sign/stroke-sets/' + setId + '/stroke/sig')
        : apiUrl('/api/sign/signers/' + sid + '/stroke/sig');
    var urlDate =
      setId && setId.length === 32
        ? apiUrl('/api/sign/stroke-sets/' + setId + '/stroke/date')
        : apiUrl('/api/sign/signers/' + sid + '/stroke/date');
    return new Promise(function (resolve, reject) {
      requestAnimationFrame(function () {
        requestAnimationFrame(function () {
          if (canvases['lib_sig_canvas'] && canvases['lib_sig_canvas'].resize) {
            canvases['lib_sig_canvas'].resize();
          }
          if (canvases['lib_date_canvas'] && canvases['lib_date_canvas'].resize) {
            canvases['lib_date_canvas'].resize();
          }
          Promise.all([
            drawUrlToCanvas('lib_sig_canvas', urlSig + ts + locParam, false),
            drawUrlToCanvas('lib_date_canvas', urlDate + ts + locParam, false),
          ])
            .then(function (results) {
              var sigOk = !!results[0];
              var dateOk = !!results[1];
              var parts = [];
              if (sigOk) parts.push('签名');
              if (dateOk) parts.push('日期');
              if (!parts.length) {
                setLibStrokeFeedback(
                  '该签署人尚无已保存的签名或日期笔迹（可先手写后保存）',
                  false
                );
              } else {
                var tail =
                  !sigOk || !dateOk
                    ? '（未载入的一侧尚无素材，可只手写该侧后点「仅保存签名」或「仅保存日期」）'
                    : '';
                setLibStrokeFeedback('已载入：' + parts.join('、') + tail, false);
              }
              resolve();
            })
            .catch(reject);
        });
      });
    });
  }

  function _libResizeStrokeCanvases() {
    if (canvases['lib_sig_canvas'] && canvases['lib_sig_canvas'].resize) {
      canvases['lib_sig_canvas'].resize();
    }
    if (canvases['lib_date_canvas'] && canvases['lib_date_canvas'].resize) {
      canvases['lib_date_canvas'].resize();
    }
  }

  /**
   * @param {'nonempty'|'sig_only'|'date_only'|'set_both'} mode
   */
  function _libStrokeFormData(mode) {
    var sigC = document.getElementById('lib_sig_canvas');
    var dateC = document.getElementById('lib_date_canvas');
    var sigBlank = isCanvasBlank(sigC);
    var dateBlank = isCanvasBlank(dateC);
    var fd = new FormData();
    if (mode === 'sig_only') {
      if (sigBlank) {
        return { error: '「仅保存签名」需要左侧签名画布有手写内容。' };
      }
      fd.append('sig', _normalizedPngDataUrl(sigC, 'sig'));
    } else if (mode === 'date_only') {
      if (dateBlank) {
        return { error: '「仅保存日期」需要右侧日期画布有手写内容。' };
      }
      fd.append('date', _normalizedPngDataUrl(dateC, 'date'));
    } else if (mode === 'set_both') {
      if (sigBlank || dateBlank) {
        return {
          error:
            '「成套保存」需要签名与日期两侧均有手写内容。只改一侧时请用「仅保存签名 / 仅保存日期」，或任一侧有内容时用「保存非空项」。',
        };
      }
      fd.append('sig', _normalizedPngDataUrl(sigC, 'sig'));
      fd.append('date', _normalizedPngDataUrl(dateC, 'date'));
    } else {
      if (sigBlank && dateBlank) {
        return {
          error:
            '请先在「签名」或「日期」画布中至少手写一项（也可用「仅保存签名 / 仅保存日期」单侧提交）。',
        };
      }
      if (!sigBlank) fd.append('sig', _normalizedPngDataUrl(sigC, 'sig'));
      if (!dateBlank) fd.append('date', _normalizedPngDataUrl(dateC, 'date'));
    }
    fd.append('locale', (libLocaleSelect && libLocaleSelect.value) ? libLocaleSelect.value : 'zh');
    return { form: fd };
  }

  function runLibStrokeSave(mode, busyBtn, busyText) {
    if (!libSignerSelect || !busyBtn) return;
    var sid = String(libActiveSignerId || '').trim();
    if (!sid) sid = getLibActiveSignerId();
    if (!sid) {
      setLibStrokeFeedback(
        '请先选择签署人（请点击「当前签署人」下方的快捷按钮，或在下拉框中选择）',
        true
      );
      return;
    }
    if (signersList.length && !_findSignerById(sid)) {
      setLibStrokeFeedback('所选签署人已不存在，请重新选择', true);
      return;
    }
    if (libSignerSelect.value && libSignerSelect.value !== sid) {
      setLibActiveSignerId(sid);
    }
    var sidFrozen = sid;
    _libResizeStrokeCanvases();
    var parsed = _libStrokeFormData(mode);
    if (parsed.error) {
      setLibStrokeFeedback(parsed.error, true);
      return;
    }
    return withButtonBusy(busyBtn, busyText || '保存中…', function () {
      var sidPut = String(libActiveSignerId || '').trim() || sidFrozen;
      if (signersList.length && !_findSignerById(sidPut)) {
        setLibStrokeFeedback('保存前签署人状态已变化，请重新点选后再保存', true);
        return Promise.resolve();
      }
      return fetchJson(apiUrl('/api/sign/signers/' + sidPut + '/strokes'), {
        method: 'PUT',
        body: parsed.form,
      }).then(function (r) {
        var jj = r.data || {};
        if (!jj.ok) {
          setLibStrokeFeedback(jj.error || '保存失败', true);
          return;
        }
        var nid = jj.stroke_set_id;
        var savedId = String(jj.signer_id || sidPut || '').trim() || sidPut;
        var signerLabel = _libSignerFeedbackLabel(savedId);
        if (jj.signer_name) {
          signerLabel = String(jj.signer_name) + '（' + (savedId.length > 10 ? savedId.slice(0, 8) + '…' : savedId) + '）';
        }
        if (savedId !== sidFrozen) {
          setLibStrokeFeedback(
            '保存目标与点选人不一致（点选 ' +
              _libSignerFeedbackLabel(sidFrozen) +
              '，实际写入 ' +
              signerLabel +
              '）',
            true
          );
          return refreshSigners();
        }
        var suffix = jj.overwritten ? '（已覆盖同内容的记录）' : '';
        var msg = '已按画布非空项保存到「' + signerLabel + '」' + suffix + '。';
        if (mode === 'sig_only') {
          msg = '已仅更新「' + signerLabel + '」的签名笔迹' + suffix + '。';
        } else if (mode === 'date_only') {
          msg = '已仅更新「' + signerLabel + '」的日期笔迹' + suffix + '。';
        } else if (mode === 'set_both') {
          msg = '已将签名与日期成套保存到「' + signerLabel + '」' + suffix + '。';
        }
        setLibStrokeFeedback(msg, false);
        return refreshSigners().then(function () {
          if (nid) {
            syncLibStrokeSetSelect();
            if (
              libStrokeSetSelect &&
              Array.prototype.some.call(libStrokeSetSelect.options, function (op) {
                return op.value === nid;
              })
            ) {
              libStrokeSetSelect.value = nid;
            }
          }
          refreshStrokeItemList();
        });
      });
    }).catch(function (e) {
      setLibStrokeFeedback(e.message || String(e), true);
    });
  }

  function showSignerListLoading() {
    if (!signerListEl) return;
    signerListEl.innerHTML = '';
    var li = document.createElement('li');
    li.textContent = '正在加载签署人列表…';
    li.style.border = 'none';
    li.style.background = 'transparent';
    signerListEl.appendChild(li);
  }

  function showFileListLoading() {
    if (!fileListEl || !listHint) return;
    fileListEl.innerHTML = '';
    listHint.style.display = 'block';
    listHint.textContent = '正在加载文件列表…';
    setFileListActionFeedback('', false);
  }

  function showSignedListLoading() {
    if (!signedListEl || !signedHint) return;
    signedListEl.innerHTML = '';
    signedHint.style.display = 'block';
    signedHint.textContent = '正在加载已签名列表…';
    setSignedListActionFeedback('', false);
  }

  function showStrokeItemsLoading() {
    if (!strokeItemListEl || !strokeItemsHint) return;
    strokeItemListEl.innerHTML = '';
    strokeItemsHint.style.display = 'block';
    strokeItemsHint.textContent = '正在加载已存储签字图片…';
    if (strokeItemsActionMsg) {
      strokeItemsActionMsg.style.display = 'none';
      strokeItemsActionMsg.textContent = '';
    }
  }

  function fileKey(f) {
    var rel =
      f && f.webkitRelativePath && String(f.webkitRelativePath).length
        ? String(f.webkitRelativePath)
        : (f && f.name) || '';
    return [rel, f && f.size, f && f.lastModified].join('|');
  }

  function _signExtOk(name) {
    var i = name.lastIndexOf('.');
    if (i < 0) return false;
    var e = name.slice(i).toLowerCase();
    return e === '.doc' || e === '.docx' || e === '.docm' ||
      e === '.xls' || e === '.xlsx' || e === '.xlsm' ||
      e === '.zip' || e === '.7z' || e === '.rar';
  }

  function filterSignFiles(fileList) {
    var out = [];
    for (var j = 0; j < fileList.length; j++) {
      var f = fileList[j];
      if (f && _signExtOk(f.name)) out.push(f);
    }
    return out;
  }

  function updatePendingHint() {
    if (!pendingSignFiles.length) {
      fileHint.textContent =
        '可多次选择并累加 .doc/.docx/.docm/.xls/.xlsx/.xlsm/.zip/.7z/.rar，或使用「选择文件夹」上传整目录（其他类型会自动忽略）';
      saveBtn.disabled = true;
      return;
    }
    fileHint.textContent = '已选 ' + pendingSignFiles.length + ' 个文件，点击下方「保存到列表」';
    saveBtn.disabled = false;
  }

  function mergePendingSignFiles(newFiles) {
    var merged = pendingSignFiles.slice();
    var seen = {};
    merged.forEach(function (f) {
      seen[fileKey(f)] = true;
    });
    newFiles.forEach(function (f) {
      var k = fileKey(f);
      if (!seen[k]) {
        seen[k] = true;
        merged.push(f);
      }
    });
    pendingSignFiles = merged;
  }

  function setupCanvas(canvas, padOpts) {
    padOpts = padOpts || {};
    // 模拟签字笔：略粗、圆角衔接、轻微墨色晕染（高 DPI 下仍用 CSS 像素坐标）
    var penLineWidth = padOpts.lineWidth != null ? padOpts.lineWidth : 4.25;
    var inkShadowBlur = padOpts.shadowBlur != null ? padOpts.shadowBlur : 0.65;
    var ctx = canvas.getContext('2d');
    // dpr 会在横竖屏/缩放时变化，resize 内会动态刷新
    var dpr = Math.min(window.devicePixelRatio || 1, 2);
    function applyPenStyle() {
      ctx.strokeStyle = '#121212';
      ctx.lineWidth = penLineWidth;
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.miterLimit = 2;
      ctx.shadowBlur = inkShadowBlur;
      ctx.shadowColor = 'rgba(0,0,0,0.28)';
      ctx.shadowOffsetX = 0;
      ctx.shadowOffsetY = 0;
    }
    function resize() {
      dpr = Math.min(window.devicePixelRatio || 1, 2);
      var rect = canvas.getBoundingClientRect();
      var w = Math.floor(rect.width * dpr);
      var h = Math.floor(rect.height * dpr);
      // 面板默认隐藏时宽高为 0，勿把 canvas 设为 0×0（会导致无法绘制且误判未签名）
      if (w < 2 || h < 2) {
        return;
      }
      if (canvas.width === w && canvas.height === h) {
        ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
        applyPenStyle();
        return;
      }
      // 改变 width/height 会清空画布；保存前若触发 resize 会擦掉笔迹导致「保存无效」
      var oldW = canvas.width;
      var oldH = canvas.height;
      var tmp = document.createElement('canvas');
      if (oldW >= 2 && oldH >= 2) {
        tmp.width = oldW;
        tmp.height = oldH;
        tmp.getContext('2d').drawImage(canvas, 0, 0);
      }
      canvas.width = w;
      canvas.height = h;
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
      applyPenStyle();
      if (tmp.width >= 2 && tmp.height >= 2) {
        ctx.save();
        ctx.setTransform(1, 0, 0, 1, 0, 0);
        // 关键：等比缩放 + 居中，避免横竖屏切换把笔迹拉伸变形
        var sx = w / oldW;
        var sy = h / oldH;
        var s = Math.min(sx, sy);
        var dw = Math.max(1, Math.floor(oldW * s));
        var dh = Math.max(1, Math.floor(oldH * s));
        var dx = Math.floor((w - dw) / 2);
        var dy = Math.floor((h - dh) / 2);
        ctx.clearRect(0, 0, w, h);
        ctx.drawImage(tmp, 0, 0, oldW, oldH, dx, dy, dw, dh);
        ctx.restore();
        applyPenStyle();
      }
    }
    resize();
    window.addEventListener('resize', resize);
    if (window.visualViewport) {
      window.visualViewport.addEventListener('resize', resize);
    }
    var drawing = false;
    function pos(e) {
      var r = canvas.getBoundingClientRect();
      var t = e.touches ? e.touches[0] : e;
      return { x: t.clientX - r.left, y: t.clientY - r.top };
    }
    function start(e) {
      e.preventDefault();
      drawing = true;
      var p = pos(e);
      applyPenStyle();
      ctx.beginPath();
      ctx.moveTo(p.x, p.y);
    }
    function move(e) {
      if (!drawing) return;
      e.preventDefault();
      var p = pos(e);
      ctx.lineTo(p.x, p.y);
      ctx.stroke();
      ctx.beginPath();
      ctx.moveTo(p.x, p.y);
    }
    function end(e) {
      if (e) e.preventDefault();
      drawing = false;
      if (isBatchWorkbenchMode() && __batchWorkbenchAdvancedOpen && selectedFileId) {
        try {
          saveCurrentFileCanvasToCache(selectedFileId);
        } catch (_) {}
      }
    }
    canvas.addEventListener('mousedown', start);
    canvas.addEventListener('mousemove', move);
    canvas.addEventListener('mouseup', end);
    canvas.addEventListener('mouseleave', end);
    canvas.addEventListener('touchstart', start, { passive: false });
    canvas.addEventListener('touchmove', move, { passive: false });
    canvas.addEventListener('touchend', end, { passive: false });
    canvas.addEventListener('touchcancel', function (e) {
      drawing = false;
      if (e) e.preventDefault();
    }, { passive: false });
    return {
      resize: resize,
      clear: function () {
        ctx.save();
        ctx.setTransform(1, 0, 0, 1, 0, 0);
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.restore();
        resize();
      },
    };
  }

  function _normalizedPngDataUrl(canvas, kind) {
    // 统一导出尺寸，避免不同设备/横竖屏导致 PNG 宽高比漂移，进而在 Word/展示中观感不一致
    if (!canvas) return '';
    kind = (kind || 'sig').toLowerCase();
    // 日期与签名使用同一目标宽度，避免入库 PNG 分辨率不同 → Word 同宽缩放时笔画粗细/墨色观感不一致
    var targetW = 1200;
    var targetH = 360;
    try {
      var srcW = canvas.width || 0;
      var srcH = canvas.height || 0;
      if (srcW < 2 || srcH < 2) {
        return canvas.toDataURL('image/png');
      }
      var out = document.createElement('canvas');
      out.width = targetW;
      out.height = targetH;
      var octx = out.getContext('2d');
      octx.setTransform(1, 0, 0, 1, 0, 0);
      // 白底：透明 PNG 在 Word 里叠在表格底纹上时，抗锯齿边缘会被看成「发灰变淡」
      octx.imageSmoothingEnabled = false;
      octx.fillStyle = '#ffffff';
      octx.fillRect(0, 0, targetW, targetH);
      var sx = targetW / srcW;
      var sy = targetH / srcH;
      var s = Math.min(sx, sy);
      var dw = Math.max(1, Math.floor(srcW * s));
      var dh = Math.max(1, Math.floor(srcH * s));
      var dx = Math.floor((targetW - dw) / 2);
      var dy = Math.floor((targetH - dh) / 2);
      octx.drawImage(canvas, 0, 0, srcW, srcH, dx, dy, dw, dh);
      return out.toDataURL('image/png');
    } catch (_) {
      try {
        return canvas.toDataURL('image/png');
      } catch (e) {
        return '';
      }
    }
  }

  function resizeCanvasesForRoles(roleIds) {
    roleIds.forEach(function (id) {
      ['sig_', 'date_'].forEach(function (prefix) {
        var key = prefix + id;
        if (canvases[key] && typeof canvases[key].resize === 'function') {
          canvases[key].resize();
        }
      });
    });
  }

  function showBatchResult(text, isErr) {
    batchResultMsg.style.display = text ? 'block' : 'none';
    batchResultMsg.style.color = isErr ? 'var(--error)' : 'var(--text-muted)';
    batchResultMsg.style.whiteSpace = text && String(text).indexOf('\n') >= 0 ? 'pre-line' : 'normal';
    batchResultMsg.textContent = text || '';
  }

  /** 与首页批量打印 ETA 文案风格一致 */
  function _formatSignBatchEtaSeconds(sec) {
    if (sec == null || typeof sec !== 'number' || !isFinite(sec) || sec < 0) return '';
    var s = Math.round(sec);
    if (s < 60) return s + ' 秒';
    var m = Math.floor(s / 60);
    var rs = s % 60;
    if (m < 60) return m + ' 分 ' + rs + ' 秒';
    var h = Math.floor(m / 60);
    var rm = m % 60;
    return h + ' 小时 ' + rm + ' 分 ' + rs + ' 秒';
  }

  function _randomSignBatchIdHex32() {
    try {
      if (typeof crypto !== 'undefined' && crypto.getRandomValues) {
        var arr = new Uint8Array(16);
        crypto.getRandomValues(arr);
        var o = '';
        for (var i = 0; i < 16; i++) {
          o += ('0' + arr[i].toString(16)).slice(-2);
        }
        return o;
      }
    } catch (_) {}
    var x = '';
    for (var j = 0; j < 32; j++) {
      x += Math.floor(Math.random() * 16).toString(16);
    }
    return x;
  }

  function drawUrlToCanvas(canvasId, url, onImgErr) {
    var canvas = document.getElementById(canvasId);
    if (!canvas) return Promise.resolve(false);
    var ctx = canvas.getContext('2d');
    return new Promise(function (resolve) {
      var img = new Image();
      if (url && String(url).indexOf('blob:') !== 0) {
        img.crossOrigin = 'anonymous';
      }
      img.onload = function () {
        try {
          if (canvases[canvasId] && canvases[canvasId].resize) {
            canvases[canvasId].resize();
          }
          var dpr = Math.min(window.devicePixelRatio || 1, 2);
          var bw = canvas.width;
          var bh = canvas.height;
          ctx.setTransform(1, 0, 0, 1, 0, 0);
          ctx.clearRect(0, 0, bw, bh);
          ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
          var cw = canvas.getBoundingClientRect().width;
          var ch = canvas.getBoundingClientRect().height;
          ctx.drawImage(img, 0, 0, cw, ch);
          resolve(true);
        } catch (_) {
          resolve(false);
        }
      };
      img.onerror = function () {
        try {
          var c2 = canvas.getContext('2d');
          if (c2) {
            c2.setTransform(1, 0, 0, 1, 0, 0);
            c2.clearRect(0, 0, canvas.width, canvas.height);
          }
        } catch (_) {}
        var msg = '无法加载手写图（请确认该签署人已保存对应签名/日期）';
        if (typeof onImgErr === 'function') {
          onImgErr(msg);
        } else if (onImgErr !== false) {
          showErr(msg);
        }
        resolve(false);
      };
      img.src = url;
    });
  }

  function _applySignersListToUi() {
    renderSignerLib();
    try {
      if (typeof window.__tryRestorePieceDraft === 'function') {
        window.__tryRestorePieceDraft();
      }
    } catch (_) {}
    if (IS_FILE_SIGN_PAGE) {
      renderNeedSignTable();
      refreshRoleSaveSignerControls();
      updateBatchUi();
      updateSubmitState();
    }
    if (IS_MATERIALS_PAGE) {
      refreshStrokeItemList();
    }
  }

  function refreshSigners(opt) {
    opt = opt || {};
    if (__refreshSignersPromise && !opt.force) {
      return __refreshSignersPromise;
    }
    if (!opt.skipPageProgress) {
      beginPageProgress(opt.progressLabel || '正在加载签署人列表…');
    }
    showSignerListLoading();
    var signersQs = '_=' + Date.now();
    if (IS_MATERIALS_PAGE) {
      signersQs += '&brief=1';
    } else if (IS_FILE_SIGN_PAGE && opt.compact !== false) {
      signersQs += '&compact=1';
    }
    var signersUrl = apiUrl('/api/sign/signers') + '?' + signersQs;
    __refreshSignersPromise = fetchJson(signersUrl, {
      cache: 'no-store',
      timeoutMs: SIGNERS_LIST_FETCH_TIMEOUT_MS,
    })
      .then(function (result) {
        var j = result.data || {};
        if (!j.ok) {
          // 软失败：保留前次 signersList，避免批量处理途中抖动一次就让后续所有文件「找不到签署人」
          if (signerListEl && !signersList.length) signerListEl.innerHTML = '';
          if (signerLibHint) {
            signerLibHint.textContent = signersList.length
              ? '签署人列表刷新失败，沿用上次结果：' + (j.error || '请稍后重试。')
              : '签署人列表加载失败：' + (j.error || '请确认服务已重启。');
          }
          if (IS_FILE_SIGN_PAGE) {
            renderNeedSignTable();
            updateBatchUi();
            updateSubmitState();
          }
          if (IS_MATERIALS_PAGE) {
            refreshStrokeItemList();
          }
          return;
        }
        signersDbShare = !!j.db_share;
        signersList = Array.isArray(j.signers) ? j.signers : [];
        if (signerLibHint) signerLibHint.textContent = signersDbShare
          ? '已启用 MySQL：签署人笔迹可在多台电脑复用。'
          : '当前为会话目录存储：笔迹仅保存在本浏览器对应会话目录，换浏览器需重新添加。';
        if (pieceDateCard) pieceDateCard.style.display = signersDbShare ? 'block' : 'none';
        if (opt.deferRender) {
          setTimeout(_applySignersListToUi, 0);
        } else {
          _applySignersListToUi();
        }
      })
      .catch(function (e) {
        // 软失败：保留前次 signersList。仅在「初次加载就失败」时清空。
        if (signerListEl && !signersList.length) signerListEl.innerHTML = '';
        if (signerLibHint) {
          signerLibHint.textContent = signersList.length
            ? '签署人列表刷新失败，沿用上次结果：' + (e && e.message ? e.message : String(e))
            : '无法加载签署人列表：' + (e && e.message ? e.message : String(e));
        }
        if (IS_FILE_SIGN_PAGE) {
          renderNeedSignTable();
          refreshRoleSaveSignerControls();
          updateBatchUi();
          updateSubmitState();
        }
        if (IS_MATERIALS_PAGE) {
          refreshStrokeItemList();
        }
      })
      .finally(function () {
        if (!opt.skipPageProgress) endPageProgress();
        if (__refreshSignersPromise) {
          __refreshSignersPromise = null;
        }
      });
    return __refreshSignersPromise;
  }

  function scheduleRefreshSigners(delayMs) {
    if (__refreshSignersDeferTimer) {
      clearTimeout(__refreshSignersDeferTimer);
    }
    __refreshSignersDeferTimer = setTimeout(function () {
      __refreshSignersDeferTimer = null;
      refreshSigners().catch(function () {});
    }, typeof delayMs === 'number' ? delayMs : 800);
  }

  function renderSignerLib() {
    if (!signerListEl) return;
    signerListEl.innerHTML = '';
    if (!signersList.length) {
      var empty = document.createElement('li');
      empty.textContent = '暂无签署人，请先添加。';
      empty.style.border = 'none';
      empty.style.background = 'transparent';
      signerListEl.appendChild(empty);
    } else {
    var total = signersList.length;
    var pageCount = Math.max(1, Math.ceil(total / signerPageSize));
    if (signerPageIndex >= pageCount) signerPageIndex = pageCount - 1;
    if (signerPageIndex < 0) signerPageIndex = 0;
    var start = signerPageIndex * signerPageSize;
    var end = Math.min(total, start + signerPageSize);
    if (signerPagerInfo) signerPagerInfo.textContent =
      '已录入 ' + total + ' 位签署人 · 第 ' + (signerPageIndex + 1) + ' / ' + pageCount + ' 页（默认仅显示 3 位）';
    if (signerPrevBtn) signerPrevBtn.disabled = signerPageIndex <= 0;
    if (signerNextBtn) signerNextBtn.disabled = signerPageIndex >= pageCount - 1;
    signersList.slice(start, end).forEach(function (s) {
      var li = document.createElement('li');
      var wrap = document.createElement('div');
      wrap.style.flex = '1';
      wrap.style.minWidth = '0';
      var t = document.createElement('div');
      t.style.fontWeight = '500';
      t.textContent = s.name || s.id;
      var meta = document.createElement('div');
      meta.className = 'signed-meta';
      var brief = signerAssetsBrief(s);
      var lineA = document.createElement('div');
      lineA.textContent = '已入库：' + brief.hint;
      var lineB = document.createElement('div');
      lineB.style.marginTop = '3px';
      lineB.style.opacity = '0.92';
      lineB.textContent = brief.statusLine;
      meta.appendChild(lineA);
      meta.appendChild(lineB);
      wrap.appendChild(t);
      wrap.appendChild(meta);
      var del = document.createElement('button');
      del.type = 'button';
      del.className = 'btn btn-secondary del-btn';
      del.textContent = '删除';
      del.addEventListener('click', function () {
        withButtonBusy(del, '删除中…', function () {
          return fetchJson(apiUrl('/api/sign/signers/' + s.id), { method: 'DELETE' }).then(
            function (r) {
              var jj = r.data;
              if (!jj.ok) {
                showSignerErr(jj.error || '删除失败');
                return;
              }
              signersList = jj.signers || [];
              showSignerErr('');
              renderSignerLib();
              renderNeedSignTable();
            }
          );
        }).catch(function (e) {
          showSignerErr(e.message || String(e));
        });
      });
      li.appendChild(wrap);
      li.appendChild(del);
      signerListEl.appendChild(li);
    });
    }
    syncLibSignerSelect();
  }

  function roleLabel(rid) {
    var x = ROLES.find(function (r) {
      return r.id === rid;
    });
    return x ? x.label : rid;
  }

  /** 合并接口返回的 roles 与 blocks 内的 role_id，提高展示/默认勾选覆盖率 */
  function mergeDetectedRolesForUi() {
    var out = [];
    var seen = {};
    if (!lastDetectData || !lastDetectData.ok) return out;
    if (selectedFileId == null) return out;
    if (String(lastDetectFileId) !== String(selectedFileId)) return out;
    (lastDetectData.roles || []).forEach(function (x) {
      if (x && x.id && !seen[x.id]) {
        seen[x.id] = true;
        out.push({ id: x.id, conf: x.confidence });
      }
    });
    (lastDetectData.blocks || []).forEach(function (b) {
      var bc = b && typeof b.confidence === 'number' ? b.confidence : null;
      (b && b.fields ? b.fields : []).forEach(function (f) {
        if (f && f.type === 'role_id' && f.name && !seen[f.name]) {
          seen[f.name] = true;
          out.push({ id: f.name, conf: bc });
        }
      });
    });
    return _filterAiwordHandoffRoles(out, lastDetectData, selectedFileId);
  }

  function _isFromAiwordHandoff() {
    try {
      return !!(document.body && document.body.classList.contains('from-aiword'));
    } catch (_) {
      return false;
    }
  }

  /** 本次 aiword 交接登记的文件 id（批量/单文件均写入 __aiwordHandoffCtxByFileId） */
  function _aiwordHandoffFileIdSet() {
    var o = __aiwordHandoffCtxByFileId || {};
    return Object.keys(o).filter(function (k) {
      return k && o[k] != null;
    });
  }

  function _isAiwordHandoffFileId(fid) {
    if (fid == null || fid === '') return false;
    return Object.prototype.hasOwnProperty.call(__aiwordHandoffCtxByFileId || {}, String(fid));
  }

  function _firstAiwordHandoffFileId() {
    var i;
    for (i = 0; i < savedFiles.length; i++) {
      var id = savedFiles[i] && savedFiles[i].id;
      if (id && _isAiwordHandoffFileId(id)) return String(id);
    }
    var ids = _aiwordHandoffFileIdSet();
    return ids.length ? String(ids[0]) : null;
  }

  function _hideAiwordHandoffLoadingMask() {
    try {
      var el = document.getElementById('aiwordHandoffLoadingMask');
      if (el) el.classList.remove('show');
    } catch (_) {}
  }

  function _setAiwordHandoffLoadingText(title, sub) {
    try {
      var t = document.getElementById('aiwordHandoffLoadingTitle');
      if (t && title) t.textContent = title;
      var s = document.getElementById('aiwordHandoffLoadingSub');
      if (s && sub) s.textContent = sub;
    } catch (_) {}
  }

  function _aiwordHandoffClaimSubtext(fileCount) {
    var n = Math.max(0, parseInt(fileCount, 10) || 0);
    if (n <= 1) {
      return '正在登记到签字列表，通常数秒即可完成，请勿关闭页面';
    }
    if (n <= 5) {
      return '正在登记 ' + n + ' 份文档到签字列表，请稍候，请勿关闭页面';
    }
    return '正在登记 ' + n + ' 份文档，数量较多时可能需 1～3 分钟，请勿关闭页面';
  }

  function _aiwordHandoffPipelineSubtext(fileCount) {
    var n = Math.max(0, parseInt(fileCount, 10) || 0);
    // 单文件 detect 受 MySQL 读 BLOB + docx 解析影响，实测每份约 5~15 秒；
    // 文案给一个保守区间，避免 1 分钟提示比实际过于乐观。
    if (n <= 1) {
      return '正在识别签字位并匹配素材，请稍候（约 5~15 秒）';
    }
    var lo = n * 5;
    var hi = n * 15;
    return (
      '正在识别并匹配 ' +
      n +
      ' 个任务（约 ' +
      lo +
      '~' +
      hi +
      ' 秒）；表格已可操作，可同时勾选/修改其它文件'
    );
  }

  function _parseBlockSourceOrder(hint) {
    // detect_fields.py:
    // - docx: "table{ti}.row{ri}" 或 "paragraph{pi}"
    // - xlsx: "{sheet}!row{ri}"
    var s = (hint == null ? '' : String(hint)).trim();
    if (!s) return null;
    var m1 = /^table(\d+)\.row(\d+)$/i.exec(s);
    if (m1) {
      return { group: 0, a: parseInt(m1[1], 10) || 0, b: parseInt(m1[2], 10) || 0 };
    }
    var m2 = /^paragraph(\d+)$/i.exec(s);
    if (m2) {
      return { group: 1, a: parseInt(m2[1], 10) || 0, b: 0 };
    }
    var m3 = /!row(\d+)$/i.exec(s);
    if (m3) {
      return { group: 0, a: 0, b: parseInt(m3[1], 10) || 0 };
    }
    return null;
  }

  function _roleMinDocOrder(rid) {
    // 返回可比较的排序键：越小越靠上；无信息则返回 null
    if (!lastDetectData || !lastDetectData.ok) return null;
    var best = null;
    (lastDetectData.blocks || []).forEach(function (b) {
      if (!b || !b.fields) return;
      var hit = false;
      (b.fields || []).forEach(function (f) {
        if (f && f.type === 'role_id' && String(f.name) === String(rid)) {
          hit = true;
        }
      });
      if (!hit) return;
      var ord = _parseBlockSourceOrder(b.source_hint);
      if (!ord) return;
      var key = [ord.group, ord.a, ord.b];
      if (!best) best = key;
      else {
        for (var i = 0; i < 3; i++) {
          if (key[i] < best[i]) {
            best = key;
            break;
          }
          if (key[i] > best[i]) break;
        }
      }
    });
    return best;
  }

  function renderNeedSignTable() {
    if (!IS_FILE_SIGN_PAGE || !needSignTable) return;
    needSignTable.innerHTML = '';
    if (!selectedFileId) {
      needSignTable.textContent = '请先选择列表中的文件。';
      return;
    }
    var roleRows = mergeDetectedRolesForUi();
    if (!roleRows.length) {
      selectedRoleIds().forEach(function (rid) {
        roleRows.push({ id: rid, conf: null });
      });
    }
    if (!roleRows.length) {
      needSignTable.textContent = lastDetectError
        ? '需签角色识别失败：' +
          lastDetectError +
          ' 请检查 FTP/网络后点「重新识别」，或在下方手动勾选「签字角色」。'
        : '未检测到需签字角色且未勾选角色。请勾选下方「签字角色」，或确认模板含编制/审核/批准等关键词后重新选择文件。';
      return;
    }
    // 关键：按文档出现顺序（从上到下）排序，避免角色列表“乱序”导致误选位置
    try {
      roleRows.sort(function (a, b) {
        var ka = _roleMinDocOrder(a.id);
        var kb = _roleMinDocOrder(b.id);
        if (ka && kb) {
          for (var i = 0; i < 3; i++) {
            if (ka[i] !== kb[i]) return ka[i] - kb[i];
          }
          return 0;
        }
        if (ka && !kb) return -1;
        if (!ka && kb) return 1;
        // 无位置时按 ROLES 默认顺序兜底
        var ia = ROLES.findIndex(function (r) { return r.id === a.id; });
        var ib = ROLES.findIndex(function (r) { return r.id === b.id; });
        if (ia < 0) ia = 9999;
        if (ib < 0) ib = 9999;
        return ia - ib;
      });
    } catch (_) {}
    var tbl = document.createElement('table');
    tbl.style.width = '100%';
    tbl.style.borderCollapse = 'collapse';
    tbl.style.fontSize = '0.9rem';
    var thead = document.createElement('thead');
    var hr = document.createElement('tr');
    ['角色', '置信度', '签名素材（签署人）', '日期素材（签署人）', '操作'].forEach(function (h) {
      var th = document.createElement('th');
      th.textContent = h;
      th.style.textAlign = 'left';
      th.style.padding = '6px 8px';
      th.style.borderBottom = '1px solid var(--border)';
      hr.appendChild(th);
    });
    thead.appendChild(hr);
    tbl.appendChild(thead);
    var tb = document.createElement('tbody');
    roleRows.forEach(function (row) {
      var rid = row.id;
      var tr = document.createElement('tr');
      var td1 = document.createElement('td');
      td1.textContent = roleLabel(rid);
      td1.style.padding = '6px 8px';
      td1.style.borderBottom = '1px solid var(--border)';
      var td2 = document.createElement('td');
      td2.style.padding = '6px 8px';
      td2.style.borderBottom = '1px solid var(--border)';
      td2.textContent =
        row.conf != null && typeof row.conf === 'number' ? row.conf.toFixed(2) : '—';
      var td3 = document.createElement('td');
      td3.style.padding = '6px 8px';
      td3.style.borderBottom = '1px solid var(--border)';
      var pair0 = currentRoleMap[rid] && typeof currentRoleMap[rid] === 'object' ? currentRoleMap[rid] : {};
      var fq = _ensureRoleFilterState(rid);
      var sigSel = document.createElement('select');
      sigSel.style.maxWidth = '100%';
      sigSel.style.padding = '6px';
      fillRoleItemSelect(sigSel, 'sig', pair0.sig || '', fq.sig);
      var sigFilter = document.createElement('input');
      sigFilter.type = 'search';
      sigFilter.placeholder = '按签署人筛选…';
      sigFilter.value = fq.sig || '';
      sigFilter.style.width = '100%';
      sigFilter.style.boxSizing = 'border-box';
      sigFilter.style.padding = '6px';
      sigFilter.style.marginBottom = '6px';
      sigFilter.style.border = '1px solid var(--border)';
      sigFilter.style.borderRadius = '8px';
      sigFilter.addEventListener('input', function () {
        fq.sig = sigFilter.value || '';
        var prev = sigSel.value || '';
        fillRoleItemSelect(sigSel, 'sig', prev, fq.sig);
      });
      sigSel.addEventListener('change', function () {
        var fid = selectedFileId;
        if (!fid) return;
        var m = Object.assign({}, currentRoleMap);
        var p = m[rid] && typeof m[rid] === 'object' ? Object.assign({}, m[rid]) : {};
        p.sig = sigSel.value || null;
        if (!roleMapEntryNonEmpty(p)) delete m[rid];
        else m[rid] = p;
        currentRoleMap = m;
        cachePatchCurrentRoleMap(fid, currentRoleMap);
        fetchJson(apiUrl('/api/sign/files/' + fid + '/role-map'), {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ map: m }),
        })
          .then(function (r) {
            var jj = r.data;
            if (String(selectedFileId) !== String(fid)) return;
            if (!jj.ok) {
              setRoleRowFeedback(rid, jj.error || '保存映射失败');
              return;
            }
            currentRoleMap = jj.map || m;
            cachePatchCurrentRoleMap(fid, currentRoleMap);
            setRoleRowFeedback(rid, '');
          })
          .catch(function (e) {
            setRoleRowFeedback(rid, e.message || String(e));
          })
          .then(function () {
            updateSubmitState();
          });

        // 选择签名后：自动把日期下拉框切换到同一签署人的日期素材（仅 item 模式时有意义）
        try {
          var opt = sigSel.options[sigSel.selectedIndex];
          var sid2 = opt ? (opt.getAttribute('data-signer-id') || '') : '';
          if (sid2) {
            // 自动筛选到该签署人，减少手动找
            fq.date = _signerNameForFilter(_findSignerById(sid2), fq.sig);
            dateFilter.value = fq.date;
            fillRoleItemSelect(dateSel, 'date', dateSel.value || '', fq.date);
            // 若当前未选日期或选的不是该签署人，自动选该签署人的第一条日期素材
            var curOpt = dateSel.options[dateSel.selectedIndex];
            var curSid = curOpt ? (curOpt.getAttribute('data-signer-id') || '') : '';
            if (!dateSel.value || (curSid && curSid !== sid2)) {
              var best = '';
              Array.prototype.some.call(dateSel.options, function (o2) {
                if (!o2 || !o2.getAttribute) return false;
                if ((o2.getAttribute('data-signer-id') || '') === sid2 && o2.value) {
                  best = o2.value;
                  return true;
                }
                return false;
              });
              if (best) {
                dateSel.value = best;
                // 触发保存 role-map
                dateSel.dispatchEvent(new Event('change'));
              }
            }
          }
        } catch (_) {}
      });
      td3.appendChild(sigFilter);
      td3.appendChild(sigSel);
      var td3b = document.createElement('td');
      td3b.style.padding = '6px 8px';
      td3b.style.borderBottom = '1px solid var(--border)';
      var dm0 = String(pair0.date_mode || 'item').toLowerCase();
      // 兼容旧值：原「英文 15.April.2026」统一改为英文空格版
      if (dm0 === 'composite_en') dm0 = 'composite_en_space';
      if (
        !signersDbShare ||
        ['composite_zh_ymd', 'composite_en_space'].indexOf(dm0) < 0
      ) {
        dm0 = 'item';
      }

      var dateItemBox = document.createElement('div');
      var dateCompBox = document.createElement('div');

      var dateSel = document.createElement('select');
      dateSel.style.maxWidth = '100%';
      dateSel.style.padding = '6px';
      fillRoleItemSelect(dateSel, 'date', pair0.date || '', fq.date);
      var dateFilter = document.createElement('input');
      dateFilter.type = 'search';
      dateFilter.placeholder = '按签署人筛选…';
      dateFilter.value = fq.date || '';
      dateFilter.style.width = '100%';
      dateFilter.style.boxSizing = 'border-box';
      dateFilter.style.padding = '6px';
      dateFilter.style.marginBottom = '6px';
      dateFilter.style.border = '1px solid var(--border)';
      dateFilter.style.borderRadius = '8px';
      dateFilter.addEventListener('input', function () {
        fq.date = dateFilter.value || '';
        var prev = dateSel.value || '';
        fillRoleItemSelect(dateSel, 'date', prev, fq.date);
      });
      dateSel.addEventListener('change', function () {
        var fid = selectedFileId;
        if (!fid) return;
        var m = Object.assign({}, currentRoleMap);
        var p = m[rid] && typeof m[rid] === 'object' ? Object.assign({}, m[rid]) : {};
        p.date = dateSel.value || null;
        if (!roleMapEntryNonEmpty(p)) delete m[rid];
        else m[rid] = p;
        currentRoleMap = m;
        cachePatchCurrentRoleMap(fid, currentRoleMap);
        fetchJson(apiUrl('/api/sign/files/' + fid + '/role-map'), {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ map: m }),
        })
          .then(function (r) {
            var jj = r.data;
            if (String(selectedFileId) !== String(fid)) return;
            if (!jj.ok) {
              setRoleRowFeedback(rid, jj.error || '保存映射失败');
              return;
            }
            currentRoleMap = jj.map || m;
            cachePatchCurrentRoleMap(fid, currentRoleMap);
            setRoleRowFeedback(rid, '');
          })
          .catch(function (e) {
            setRoleRowFeedback(rid, e.message || String(e));
          })
          .then(function () {
            updateSubmitState();
          });
      });
      dateItemBox.appendChild(dateFilter);
      dateItemBox.appendChild(dateSel);

      var dateModeSel = document.createElement('select');
      dateModeSel.style.width = '100%';
      dateModeSel.style.boxSizing = 'border-box';
      dateModeSel.style.padding = '6px';
      dateModeSel.style.marginBottom = '8px';
      dateModeSel.style.border = '1px solid var(--border)';
      dateModeSel.style.borderRadius = '8px';
      dateModeSel.style.fontSize = '0.88rem';
      ;[
        { v: 'item', t: '整张日期手写图' },
        { v: 'composite_zh_ymd', t: '拼接：中文 2026.04.15' },
        { v: 'composite_en_space', t: '拼接：英文 15 April 2026' },
      ].forEach(function (opt) {
        var o = document.createElement('option');
        o.value = opt.v;
        o.textContent = opt.t;
        dateModeSel.appendChild(o);
      });
      dateModeSel.value = dm0;

      var dateIsoInp = document.createElement('input');
      dateIsoInp.type = 'date';
      dateIsoInp.style.width = '100%';
      dateIsoInp.style.boxSizing = 'border-box';
      dateIsoInp.style.padding = '6px';
      dateIsoInp.style.marginBottom = '6px';
      dateIsoInp.style.border = '1px solid var(--border)';
      dateIsoInp.style.borderRadius = '8px';
      dateIsoInp.value = _singleFileCompositeDateIso(pair0);

      var pvBtn = document.createElement('button');
      pvBtn.type = 'button';
      pvBtn.className = 'btn btn-secondary';
      pvBtn.style.marginBottom = '6px';
      pvBtn.textContent = '预览拼接';
      pvBtn.title = '根据所选签署人（签名素材）与日期，预览横向拼接效果';

      var compHint = document.createElement('div');
      compHint.className = 'hint';
      compHint.style.marginTop = '4px';
      compHint.textContent = '';

      function applyDateModeVisibility(mode) {
        var comp = mode && mode !== 'item';
        dateItemBox.style.display = comp ? 'none' : 'block';
        dateCompBox.style.display = comp ? 'block' : 'none';
        var d = String(mode || '').toLowerCase();
        if (d === 'composite_zh_ymd') {
          compHint.textContent =
            '需该签署人已录入数字 0–9 与句点「.」笔迹；按 年.月.日 拼接。';
        } else if (d === 'composite_en_space') {
          compHint.textContent =
            '需已录入日数字、英文月份简称、年数字；日与月、月与年之间用空格（无需空格笔迹）。';
        } else {
          compHint.textContent = '';
        }
      }

      function putRoleMap(p) {
        var fid = selectedFileId;
        if (!fid) {
          return Promise.reject(new Error('no_file'));
        }
        var m = Object.assign({}, currentRoleMap);
        if (!roleMapEntryNonEmpty(p)) delete m[rid];
        else m[rid] = p;
        currentRoleMap = m;
        cachePatchCurrentRoleMap(fid, currentRoleMap);
        return fetchJson(apiUrl('/api/sign/files/' + fid + '/role-map'), {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ map: m }),
        })
          .then(function (r) {
            var jj = r.data;
            if (String(selectedFileId) !== String(fid)) return;
            if (!jj.ok) {
              setRoleRowFeedback(rid, jj.error || '保存映射失败');
              return;
            }
            currentRoleMap = jj.map || m;
            cachePatchCurrentRoleMap(fid, currentRoleMap);
            setRoleRowFeedback(rid, '');
          })
          .catch(function (e) {
            setRoleRowFeedback(rid, e.message || String(e));
          })
          .then(function () {
            updateSubmitState();
          });
      }

      dateCompBox.appendChild(dateIsoInp);
      dateCompBox.appendChild(pvBtn);
      dateCompBox.appendChild(compHint);
      var pvMsg = document.createElement('div');
      pvMsg.className = 'hint';
      pvMsg.style.marginTop = '6px';
      pvMsg.style.display = 'none';
      var pvImg = document.createElement('img');
      pvImg.style.display = 'none';
      pvImg.style.maxWidth = '100%';
      pvImg.style.border = '1px solid var(--border)';
      pvImg.style.borderRadius = '8px';
      pvImg.style.background = '#fff';
      pvImg.style.marginTop = '6px';
      pvImg.alt = '拼接预览';
      dateCompBox.appendChild(pvMsg);
      dateCompBox.appendChild(pvImg);

      pvBtn.addEventListener('click', function () {
        if (!signersDbShare) return;
        var m = Object.assign({}, currentRoleMap);
        var p = m[rid] && typeof m[rid] === 'object' ? Object.assign({}, m[rid]) : {};
        var sid = '';
        // 优先从当前行选中的「签名素材」推断签署人；否则用上方“当前签署人”
        try {
          var optSig = sigSel.options[sigSel.selectedIndex];
          sid = optSig ? (optSig.getAttribute('data-signer-id') || '') : '';
        } catch (_) {}
        if (!sid && libSignerSelect && libSignerSelect.value) sid = libSignerSelect.value;
        if (!sid) {
          setRoleRowFeedback(rid, '请先选择签署人（或先选择签名素材以自动确定签署人）');
          return;
        }
        var iso = (dateIsoInp.value || '').trim();
        if (!iso) {
          setRoleRowFeedback(rid, '请先选择日期');
          return;
        }
        var lay = compositeModeToLayout(dateModeSel.value);
        var u =
          apiUrl('/api/sign/signers/' + sid + '/composite-date-preview') +
          '?iso=' +
          encodeURIComponent(iso) +
          '&layout=' +
          encodeURIComponent(lay) +
          '&_=' +
          Date.now();
        pvMsg.style.display = 'none';
        pvMsg.textContent = '';
        pvImg.style.display = 'none';
        pvImg.removeAttribute('src');
        withButtonBusy(pvBtn, '预览中…', function () {
          return fetch(u, { credentials: 'include' }).then(function (res) {
            if (res.ok) {
              return res.blob().then(function (blob) {
                var url = '';
                try { url = URL.createObjectURL(blob); } catch (_) { url = ''; }
                if (!url) {
                  pvMsg.style.display = 'block';
                  pvMsg.style.color = 'var(--error)';
                  pvMsg.textContent = '预览图片生成成功，但无法在浏览器内显示（URL.createObjectURL 失败）。';
                  return;
                }
                pvImg.onload = function () {
                  try { URL.revokeObjectURL(url); } catch (_) {}
                };
                pvImg.src = url;
                pvImg.style.display = 'block';
                pvMsg.style.display = 'none';
                pvMsg.style.whiteSpace = '';
              });
            }
            return res.text().then(function (text) {
              var msg = '预览失败（HTTP ' + res.status + '）。';
              var t = (text || '').trim();
              if (t && t.charAt(0) !== '<') {
                try {
                  var j = JSON.parse(text);
                  if (j && j.error) msg = String(j.error);
                } catch (_) {}
              }
              pvMsg.style.display = 'block';
              pvMsg.style.color = 'var(--error)';
              pvMsg.style.whiteSpace = 'pre-wrap';
              pvMsg.textContent = msg;
            });
          });
        }).catch(function (e) {
          pvMsg.style.display = 'block';
          pvMsg.style.color = 'var(--error)';
          pvMsg.style.whiteSpace = 'pre-wrap';
          pvMsg.textContent = (e && e.message) ? e.message : String(e);
        });
      });

      dateIsoInp.addEventListener('change', function () {
        var m = Object.assign({}, currentRoleMap);
        var p = m[rid] && typeof m[rid] === 'object' ? Object.assign({}, m[rid]) : {};
        var v = dateModeSel.value || 'item';
        if (v === 'item') return;
        p.date_mode = v;
        p.date = null;
        p.date_iso = dateIsoInp.value || null;
        putRoleMap(p).catch(function (e) {
          setRoleRowFeedback(rid, e.message || String(e));
        });
      });

      function syncModeFromSelect() {
        var v = dateModeSel.value || 'item';
        var m = Object.assign({}, currentRoleMap);
        var p = m[rid] && typeof m[rid] === 'object' ? Object.assign({}, m[rid]) : {};
        if (v !== 'item') {
          p.date_mode = v;
          p.date = null;
          p.date_iso = dateIsoInp.value || _singleFileCompositeDateIso(p) || null;
        } else {
          p.date_mode = 'item';
          p.date_iso = null;
        }
        applyDateModeVisibility(v);
        putRoleMap(p).catch(function (e) {
          setRoleRowFeedback(rid, e.message || String(e));
        });
      }
      dateModeSel.addEventListener('change', syncModeFromSelect);

      if (signersDbShare) {
        var dmLbl = document.createElement('div');
        dmLbl.style.fontSize = '0.82rem';
        dmLbl.style.color = 'var(--text-muted)';
        dmLbl.style.marginBottom = '4px';
        dmLbl.textContent = '日期方式';
        td3b.appendChild(dmLbl);
        td3b.appendChild(dateModeSel);
        applyDateModeVisibility(dm0);
      } else {
        dateCompBox.style.display = 'none';
        applyDateModeVisibility('item');
      }

      td3b.appendChild(dateItemBox);
      td3b.appendChild(dateCompBox);
      var td4 = document.createElement('td');
      td4.style.padding = '6px 8px';
      td4.style.borderBottom = '1px solid var(--border)';
      td4.style.whiteSpace = 'normal';
      var locSel = document.createElement('select');
      locSel.style.padding = '6px';
      locSel.style.marginRight = '6px';
      locSel.title = '该角色入库时使用中文/英文版本（可自动匹配）';
      var lo0 = document.createElement('option');
      lo0.value = 'auto';
      lo0.textContent = '自动';
      var lo1 = document.createElement('option');
      lo1.value = 'zh';
      lo1.textContent = '中文';
      var lo2 = document.createElement('option');
      lo2.value = 'en';
      lo2.textContent = '英文';
      locSel.appendChild(lo0);
      locSel.appendChild(lo1);
      locSel.appendChild(lo2);
      locSel.value = roleLocaleMap[rid] || 'auto';
      locSel.addEventListener('change', function () {
        roleLocaleMap[rid] = locSel.value || 'auto';
        roleLocaleManual[rid] = roleLocaleMap[rid] !== 'auto';
      });
      var bLoad = document.createElement('button');
      bLoad.type = 'button';
      bLoad.className = 'btn btn-secondary';
      bLoad.textContent = '载入签名';
      bLoad.style.marginRight = '6px';
      bLoad.addEventListener('click', function () {
        var itemId = sigSel.value;
        if (!itemId) {
          setRoleRowFeedback(rid, '请先选择签名素材');
          return;
        }
        withButtonBusy(bLoad, '载入中…', function () {
          return new Promise(function (resolve) {
            setRoleRowFeedback(rid, '');
            var ts = '?t=' + Date.now();
            setRoleChecked(rid, true);
            requestAnimationFrame(function () {
              requestAnimationFrame(function () {
                resizeCanvasesForRoles([rid]);
                drawUrlToCanvas(
                  'sig_' + rid,
                  apiUrl('/api/sign/stroke-items/' + itemId + '/png') + ts,
                  function (msg) {
                    setRoleRowFeedback(rid, msg);
                  }
                ).then(function () {
                  resolve();
                });
              });
            });
          });
        }).then(function () {
          updateSubmitState();
        });
      });
      var bLoadDate = document.createElement('button');
      bLoadDate.type = 'button';
      bLoadDate.className = 'btn btn-secondary';
      bLoadDate.textContent = '载入日期';
      bLoadDate.style.marginRight = '6px';
      bLoadDate.addEventListener('click', function () {
        var pr = currentRoleMap[rid];
        var dmRow = (dateModeSel && dateModeSel.value) || (pr && pr.date_mode) || 'item';
        if (isCompositeDateMode(dmRow)) {
          if (!signersDbShare) {
            setRoleRowFeedback(rid, '拼接日期需启用 MySQL。');
            return;
          }
          var sidC = '';
          try {
            var optSigC = sigSel.options[sigSel.selectedIndex];
            sidC = optSigC ? (optSigC.getAttribute('data-signer-id') || '') : '';
          } catch (_) {}
          if (!sidC && libSignerSelect && libSignerSelect.value) sidC = libSignerSelect.value;
          if (!sidC) {
            setRoleRowFeedback(
              rid,
              '请先选择签名素材（用于确定签署人），或在上方的「当前签署人」中选择。'
            );
            return;
          }
          var isoC = (dateIsoInp.value || '').trim();
          if (!isoC) {
            setRoleRowFeedback(rid, '请先选择日历日期');
            return;
          }
          var layC = compositeModeToLayout(dmRow);
          var uC =
            apiUrl('/api/sign/signers/' + sidC + '/composite-date-preview') +
            '?iso=' +
            encodeURIComponent(isoC) +
            '&layout=' +
            encodeURIComponent(layC) +
            '&_=' +
            Date.now();
          withButtonBusy(bLoadDate, '载入中…', function () {
            return fetch(uC, { credentials: 'include' }).then(function (res) {
              if (!res.ok) {
                return res.text().then(function (text) {
                  var msg = '载入拼接日期失败（HTTP ' + res.status + '）';
                  var t = (text || '').trim();
                  if (t && t.charAt(0) !== '<') {
                    try {
                      var jce = JSON.parse(t);
                      if (jce && jce.error) msg = String(jce.error);
                    } catch (_) {}
                  }
                  setRoleRowFeedback(rid, msg);
                  return Promise.reject(new Error(msg));
                });
              }
              return res.blob().then(function (blob) {
                var ourl = '';
                try {
                  ourl = URL.createObjectURL(blob);
                } catch (e0) {
                  setRoleRowFeedback(rid, '无法在浏览器中生成本地图片地址');
                  return Promise.reject(e0);
                }
                setRoleRowFeedback(rid, '');
                setRoleChecked(rid, true);
                return new Promise(function (resolve) {
                  requestAnimationFrame(function () {
                    requestAnimationFrame(function () {
                      resizeCanvasesForRoles([rid]);
                      drawUrlToCanvas('date_' + rid, ourl, function (m) {
                        setRoleRowFeedback(rid, m || '拼接图无法绘制到画布');
                      }).then(function () {
                        try {
                          URL.revokeObjectURL(ourl);
                        } catch (_) {}
                        resolve();
                      });
                    });
                  });
                });
              });
            });
          }).then(function () {
            updateSubmitState();
          });
          return;
        }
        var itemId = dateSel.value;
        if (!itemId) {
          setRoleRowFeedback(rid, '请先选择日期素材');
          return;
        }
        withButtonBusy(bLoadDate, '载入中…', function () {
          return new Promise(function (resolve) {
            setRoleRowFeedback(rid, '');
            var ts = '?t=' + Date.now();
            setRoleChecked(rid, true);
            requestAnimationFrame(function () {
              requestAnimationFrame(function () {
                resizeCanvasesForRoles([rid]);
                drawUrlToCanvas(
                  'date_' + rid,
                  apiUrl('/api/sign/stroke-items/' + itemId + '/png') + ts,
                  function (msg) {
                    setRoleRowFeedback(rid, msg);
                  }
                ).then(function () {
                  resolve();
                });
              });
            });
          });
        }).then(function () {
          updateSubmitState();
        });
      });

      td4.appendChild(locSel);
      td4.appendChild(bLoad);
      td4.appendChild(bLoadDate);
      var rowInlineMsg = document.createElement('p');
      rowInlineMsg.id = 'needSignRowMsg_' + rid;
      rowInlineMsg.className = 'btn-inline-feedback';
      rowInlineMsg.style.display = 'none';
      rowInlineMsg.style.marginTop = '8px';
      rowInlineMsg.style.maxWidth = '100%';
      rowInlineMsg.setAttribute('role', 'status');
      td4.appendChild(rowInlineMsg);
      tr.appendChild(td1);
      tr.appendChild(td2);
      tr.appendChild(td3);
      tr.appendChild(td3b);
      tr.appendChild(td4);
      tb.appendChild(tr);
    });
    tbl.appendChild(tb);
    needSignTable.appendChild(tbl);
  }

  function updateBatchUi() {
    var n = _checkedBatchFileIds().length;
    // 批量模式依赖 MySQL；禁用时自动退出批量
    if (batchModeCb) {
      batchModeCb.disabled = !signersDbShare;
      if (!signersDbShare) batchModeCb.checked = false;
    }
    updateSubmitState();
  }

  if (batchSelectAll) {
    batchSelectAll.addEventListener('change', function () {
      document.querySelectorAll('.batch-pick').forEach(function (cb) {
        cb.checked = batchSelectAll.checked;
      });
      updateBatchUi();
    });
  }

  if (batchWorkbenchHeadCheck) {
    batchWorkbenchHeadCheck.addEventListener('change', function () {
      var on = !!batchWorkbenchHeadCheck.checked;
      _setWorkbenchRowsSelected(on, _workbenchHasActiveFilter());
      if (batchWorkbenchSelectAll) batchWorkbenchSelectAll.checked = on;
      syncHiddenBatchPicks();
      renderBatchWorkbenchTable();
    });
  }
  if (batchWorkbenchSelectAll) {
    batchWorkbenchSelectAll.addEventListener('change', function () {
      var on = !!batchWorkbenchSelectAll.checked;
      if (batchWorkbenchHeadCheck) batchWorkbenchHeadCheck.checked = on;
      _setWorkbenchRowsSelected(on, _workbenchHasActiveFilter());
      syncHiddenBatchPicks();
      renderBatchWorkbenchTable();
    });
  }
  if (batchWorkbenchSelectFilteredBtn) {
    batchWorkbenchSelectFilteredBtn.addEventListener('click', function () {
      var n = selectWorkbenchFilteredRows(true, { merge: true });
      var totalSel = _countWorkbenchSelectedRows();
      setBatchWorkbenchMsg(
        n
          ? '本次累加勾选 ' + n + ' 个（合计已选 ' + totalSel + ' 个，含筛选外仍保留）'
          : '当前筛选下没有可见文件',
        !n
      );
    });
  }
  if (batchWorkbenchDetectBtn) {
    batchWorkbenchDetectBtn.addEventListener('click', function () {
      withButtonBusy(batchWorkbenchDetectBtn, '识别中…', function () {
        return processBatchWorkbenchSelected(true, { forceReprocess: true });
      }, { skipPageProgress: true });
    });
  }
  if (batchWorkbenchApplyBtn) {
    batchWorkbenchApplyBtn.addEventListener('click', function () {
      withButtonBusy(batchWorkbenchApplyBtn, '匹配中…', function () {
        return processBatchWorkbenchSelected(false, {
          skipDetectIfCached: true,
          forceReprocess: true,
        });
      }, { skipPageProgress: true });
    });
  }
  var batchWorkbenchMarkDetectWrongBtn = document.getElementById(
    'batchWorkbenchMarkDetectWrongBtn'
  );
  if (batchWorkbenchMarkDetectWrongBtn) {
    batchWorkbenchMarkDetectWrongBtn.addEventListener('click', function () {
      markWorkbenchSelectedAsDetectWrong();
    });
  }
  initDetectCorrectionModal();
  if (batchWorkbenchDeleteBtn) {
    batchWorkbenchDeleteBtn.addEventListener('click', function () {
      withButtonBusy(batchWorkbenchDeleteBtn, '删除中…', function () {
        return batchDeleteSelectedFiles('workbench');
      }, { skipPageProgress: true });
    });
  }
  if (batchWorkbenchLocaleApplyBtn) {
    batchWorkbenchLocaleApplyBtn.addEventListener('click', function () {
      var loc = batchWorkbenchLocaleBulk ? String(batchWorkbenchLocaleBulk.value || '') : '';
      if (!loc) {
        setBatchWorkbenchMsg('请选择要批量设置的版本', true);
        return;
      }
      savedFiles.forEach(function (r) {
        if (!r || !r.id) return;
        var row = __batchWorkbenchRows[String(r.id)];
        if (row && row.selected) {
          row.locale = loc;
          syncRowToHandoffCtx(String(r.id));
        }
      });
      renderBatchWorkbenchTable();
      setBatchWorkbenchMsg('已批量设置签字版本为' + (loc === 'en' ? '英文' : '中文'), false);
    });
  }
  if (batchWorkbenchExportIssuesBtn) {
    batchWorkbenchExportIssuesBtn.addEventListener('click', function () {
      exportBatchWorkbenchIssues();
    });
  }
  if (batchWorkbenchSignBtn) {
    batchWorkbenchSignBtn.addEventListener('click', function () {
      withButtonBusy(batchWorkbenchSignBtn, '签字中…', function () {
        syncHiddenBatchPicks();
        if (batchModeCb) batchModeCb.checked = true;
        if (signSourceMode) signSourceMode.value = 'library';
        saveCurrentFileCanvasToCache(selectedFileId);
        return _doBatchSignFromSubmit({
          apply_person: true,
          apply_date: true,
          workbenchMode: true,
          feedbackTarget: 'workbench',
        });
      }, { skipPageProgress: true, skipRestoreDisabled: true });
    });
  }
  if (batchWorkbenchAdvancedCb) {
    batchWorkbenchAdvancedCb.addEventListener('change', function () {
      setBatchWorkbenchAdvanced(!!batchWorkbenchAdvancedCb.checked);
    });
  }

  if (batchModeCb) {
    batchModeCb.addEventListener('change', function () {
      showErr('');
      updateSubmitState();
    });
  }

  if (signSourceMode) {
    signSourceMode.addEventListener('change', function () {
      syncLibraryRolesModeRow();
      updateSubmitState();
    });
  }
  if (libraryRolesUseChecksCb) {
    libraryRolesUseChecksCb.addEventListener('change', updateSubmitState);
  }

  if (addSignerBtn && newSignerName) addSignerBtn.addEventListener('click', function () {
    var raw = newSignerName.value.trim();
    if (!raw) {
      showSignerErr('请填写至少一个签署人名称');
      return;
    }
    var names = parseSignerNamesInput(raw);
    if (!names.length) {
      showSignerErr('未能解析出有效姓名，请用逗号、分号或换行分隔');
      return;
    }
    withButtonBusy(addSignerBtn, '添加中…', function () {
      return fetchJson(apiUrl('/api/sign/signers'), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ names: names }),
      }).then(function (r) {
        var j = r.data || {};
        if (!j.ok) {
          showSignerErr(j.error || '添加失败');
          return;
        }
        newSignerName.value = '';
        signersList = j.signers || [];
        signersDbShare = !!j.db_share;
        var n = typeof j.added === 'number' ? j.added : names.length;
        renderSignerLib();
        renderNeedSignTable();
        showSignerErr('已添加 ' + n + ' 位签署人，可在下方为其录入手写图。');
      });
    }).catch(function (e) {
      showSignerErr(e.message || String(e));
    });
  });

  if (libSignerSelect) libSignerSelect.addEventListener('change', function () {
    showSignerErr('');
    setLibActiveSignerId(libSignerSelect.value || '');
  });

  if (libSignerFilter) libSignerFilter.addEventListener('input', function () {
    showSignerErr('');
    syncLibSignerSelect();
    syncCurrentSignerBanners();
    renderLibSignerQuickPick();
  });
  if (libLocaleSelect) libLocaleSelect.addEventListener('change', function () {
    renderLibLocaleQuickPick();
  });

  if (signerPrevBtn) signerPrevBtn.addEventListener('click', function () {
    signerPageIndex = Math.max(0, signerPageIndex - 1);
    renderSignerLib();
  });

  if (signerNextBtn) signerNextBtn.addEventListener('click', function () {
    signerPageIndex = signerPageIndex + 1;
    renderSignerLib();
  });

  if (libClearSigBtn) libClearSigBtn.addEventListener('click', function () {
    if (canvases['lib_sig_canvas'] && canvases['lib_sig_canvas'].clear) {
      canvases['lib_sig_canvas'].clear();
    }
  });

  if (libClearDateBtn) libClearDateBtn.addEventListener('click', function () {
    if (canvases['lib_date_canvas'] && canvases['lib_date_canvas'].clear) {
      canvases['lib_date_canvas'].clear();
    }
  });

  if (libLoadStrokesBtn) libLoadStrokesBtn.addEventListener('click', function () {
    if (!getLibActiveSignerId()) {
      setLibStrokeFeedback('请先在「当前签署人」中选择一位（可点快捷按钮）', true);
      return;
    }
    withButtonBusy(libLoadStrokesBtn, '载入中…', function () {
      return loadLibStrokesFromServer().catch(function (e) {
        if (e && e.message === 'no_signer') return;
        setLibStrokeFeedback(e.message || String(e), true);
      });
    });
  });

  if (libSaveSigOnlyBtn) {
    libSaveSigOnlyBtn.addEventListener('click', function () {
      runLibStrokeSave('sig_only', libSaveSigOnlyBtn, '保存签名…');
    });
  }
  if (libSaveDateOnlyBtn) {
    libSaveDateOnlyBtn.addEventListener('click', function () {
      runLibStrokeSave('date_only', libSaveDateOnlyBtn, '保存日期…');
    });
  }
  if (libSaveStrokesBtn) {
    libSaveStrokesBtn.addEventListener('click', function () {
      runLibStrokeSave('nonempty', libSaveStrokesBtn, '保存中…');
    });
  }
  if (libSaveStrokeSetBtn) {
    libSaveStrokeSetBtn.addEventListener('click', function () {
      runLibStrokeSave('set_both', libSaveStrokeSetBtn, '成套保存中…');
    });
  }

  if (btnRefreshSigners) {
    btnRefreshSigners.addEventListener('click', function () {
      withButtonBusy(btnRefreshSigners, '刷新中…', function () {
        return refreshSigners({ force: true });
      });
    });
  }
  if (btnRefreshStrokeItems) {
    btnRefreshStrokeItems.addEventListener('click', function () {
      withButtonBusy(btnRefreshStrokeItems, '刷新中…', function () {
        return refreshStrokeItemList({ force: true });
      });
    });
  }
  if (batchWorkbenchRefreshBtn) {
    batchWorkbenchRefreshBtn.addEventListener('click', function () {
      withButtonBusy(batchWorkbenchRefreshBtn, '刷新中…', function () {
        return refreshWorkbenchFilesFromServer({ silent: false });
      });
    });
  }
  if (batchWorkbenchFilterName) {
    batchWorkbenchFilterName.addEventListener('input', function () {
      __wbFilterNameDraft = String(batchWorkbenchFilterName.value || '');
    });
    batchWorkbenchFilterName.addEventListener('keydown', function (ev) {
      if (ev.key === 'Enter') {
        ev.preventDefault();
        if (_pinWorkbenchFilterNameTerm()) {
          setBatchWorkbenchMsg(
            '已筛选文件名：' + __wbFilterNameTerms.join('、') + '（回车可继续追加，OR 合并）',
            false
          );
        } else {
          setBatchWorkbenchMsg('请输入文件名关键词后回车', true);
        }
      }
    });
  }
  if (batchWorkbenchSelectByStatusBtn) {
    batchWorkbenchSelectByStatusBtn.addEventListener('click', function () {
      var n = selectWorkbenchRowsByStatus(true, { merge: true });
      var totalSel = _countWorkbenchSelectedRows();
      setBatchWorkbenchMsg(
        n
          ? '已按状态累加勾选 ' + n + ' 个（合计已选 ' + totalSel + ' 个）'
          : '请先勾选状态，再点「按状态全选」',
        !n
      );
    });
  }
  if (batchWorkbenchFilterClearBtn) {
    batchWorkbenchFilterClearBtn.addEventListener('click', function () {
      _clearWorkbenchFilters();
      setBatchWorkbenchMsg('已清除筛选条件', false);
    });
  }
  var batchWorkbenchFilterDetectIssuesBtn = document.getElementById(
    'batchWorkbenchFilterDetectIssuesBtn'
  );
  _bindWorkbenchStatusFilterCheckboxes();
  _refreshWorkbenchSlotFilterOptions();
  if (batchWorkbenchFilterStatusToggle) {
    batchWorkbenchFilterStatusToggle.addEventListener('click', function (ev) {
      ev.stopPropagation();
      var open = batchWorkbenchFilterStatusWrap &&
        batchWorkbenchFilterStatusWrap.classList.contains('open');
      _toggleWorkbenchSlotFilterPanel(false);
      _toggleWorkbenchStatusFilterPanel(!open);
    });
  }
  if (batchWorkbenchFilterSlotToggle) {
    batchWorkbenchFilterSlotToggle.addEventListener('click', function (ev) {
      ev.stopPropagation();
      var open = batchWorkbenchFilterSlotWrap &&
        batchWorkbenchFilterSlotWrap.classList.contains('open');
      _toggleWorkbenchStatusFilterPanel(false);
      _toggleWorkbenchSlotFilterPanel(!open);
    });
  }
  document.addEventListener('click', function (ev) {
    if (
      (batchWorkbenchFilterStatusWrap && batchWorkbenchFilterStatusWrap.contains(ev.target)) ||
      (batchWorkbenchFilterSlotWrap && batchWorkbenchFilterSlotWrap.contains(ev.target))
    ) {
      return;
    }
    _toggleWorkbenchStatusFilterPanel(false);
    _toggleWorkbenchSlotFilterPanel(false);
  });
  if (batchWorkbenchStatusSelectAllBtn) {
    batchWorkbenchStatusSelectAllBtn.addEventListener('click', function () {
      _selectAllWorkbenchStatusFilters();
      setBatchWorkbenchMsg('已全选状态（显示全部文件）', false);
    });
  }
  if (batchWorkbenchStatusClearBtn) {
    batchWorkbenchStatusClearBtn.addEventListener('click', function () {
      _clearWorkbenchStatusFilters();
      setBatchWorkbenchMsg('已清空状态筛选', false);
    });
  }
  if (batchWorkbenchSlotSelectAllBtn) {
    batchWorkbenchSlotSelectAllBtn.addEventListener('click', function () {
      _selectAllWorkbenchSlotFilters();
      setBatchWorkbenchMsg('已全选签字位标签', false);
    });
  }
  if (batchWorkbenchSlotClearBtn) {
    batchWorkbenchSlotClearBtn.addEventListener('click', function () {
      _clearWorkbenchSlotFilters();
      setBatchWorkbenchMsg('已清空签字位标签筛选', false);
    });
  }
  if (batchWorkbenchFilterDetectIssuesBtn) {
    batchWorkbenchFilterDetectIssuesBtn.addEventListener('click', function () {
      _setWorkbenchFilterDetectIssuesOnly();
      setBatchWorkbenchMsg(
        '已筛选：识别有误(人工) / 失败 / 超时 / 未识别。可「累加勾选当前筛选」后改规则并批量重新识别。',
        false
      );
    });
  }
  if (batchWorkbenchFilterNotSignableBtn) {
    batchWorkbenchFilterNotSignableBtn.addEventListener('click', function () {
      _setWorkbenchFilterNotSignableOnly();
      setBatchWorkbenchMsg('已筛选：仅显示不可签文件（签字位可落位校验未通过）。', false);
    });
  }
  if (btnRefreshSigned) {
    btnRefreshSigned.addEventListener('click', function () {
      withButtonBusy(btnRefreshSigned, '刷新中…', function () {
        return refreshSignedList({ force: true });
      });
    });
  }

  function _setBatchPickByIds(ids) {
    var m = {};
    (ids || []).forEach(function (x) {
      m[String(x)] = true;
    });
    document.querySelectorAll('.batch-pick').forEach(function (cb) {
      var id = String(cb.getAttribute('data-id') || '');
      cb.checked = !!m[id];
    });
    updateBatchUi();
  }

  function runAiwordGroupedBatchSign() {
    if (!aiwordBatchRunBtn) return Promise.resolve();
    var groups = computeAiwordBatchGroups();
    if (!groups.length) {
      setAiwordBatchStrategyMsg('当前无可执行分组。', true);
      return Promise.resolve();
    }
    return runWithPageProgress('按策略批量签字（共 ' + groups.length + ' 组）…', function () {
    var mode = aiwordBatchExecMode ? String(aiwordBatchExecMode.value || 'full') : 'full';
    var chain = Promise.resolve();
    aiwordBatchRunBtn.disabled = true;
    groups.forEach(function (g, idx) {
      chain = chain.then(function () {
        _setBatchPickByIds(g.file_ids || []);
        updatePageProgress('按策略批量签字 第 ' + (idx + 1) + '/' + groups.length + ' 组…', {
          done: idx,
          total: groups.length,
        });
        setAiwordBatchStrategyMsg('正在执行第 ' + (idx + 1) + '/' + groups.length + ' 组…', false);
        var signOpt = { skipPageProgress: true };
        if (mode === 'person_only') {
          return _doBatchSignFromSubmit(
            Object.assign({ apply_person: true, apply_date: false }, signOpt)
          );
        }
        if (mode === 'date_only') {
          return _doBatchSignFromSubmit(
            Object.assign({ apply_person: false, apply_date: true }, signOpt)
          );
        }
        if (mode === 'person_then_date') {
          return _doBatchSignFromSubmit(
            Object.assign({ apply_person: true, apply_date: false }, signOpt)
          ).then(function () {
            return _doBatchSignFromSubmit(
              Object.assign({ apply_person: false, apply_date: true }, signOpt)
            );
          });
        }
        return _doBatchSignFromSubmit(
          Object.assign({ apply_person: true, apply_date: true }, signOpt)
        );
      });
    });
    return chain
      .then(function () {
        setAiwordBatchStrategyMsg('已按策略完成 ' + groups.length + ' 组批量签字。', false);
      })
      .catch(function (e) {
        setAiwordBatchStrategyMsg((e && e.message) || String(e) || '执行失败', true);
      })
      .then(function () {
        aiwordBatchRunBtn.disabled = false;
      });
    });
  }

  if (aiwordBatchPreviewBtn) {
    aiwordBatchPreviewBtn.addEventListener('click', function () {
      renderAiwordBatchGroupsPreview();
    });
  }
  if (aiwordBatchApplySelectBtn) {
    aiwordBatchApplySelectBtn.addEventListener('click', function () {
      applyAiwordBatchSelectionByGroups();
    });
  }
  if (aiwordBatchExportBtn) {
    aiwordBatchExportBtn.addEventListener('click', function () {
      exportAiwordBatchGroupsSnapshot();
    });
  }
  if (aiwordBatchRunBtn) {
    aiwordBatchRunBtn.addEventListener('click', function () {
      withButtonBusy(aiwordBatchRunBtn, '执行中…', function () {
        return runAiwordGroupedBatchSign();
      }, { skipPageProgress: true });
    });
  }
  if (aiwordBatchGroupStrategy) {
    aiwordBatchGroupStrategy.addEventListener('change', function () {
      renderAiwordBatchGroupsPreview();
    });
  }

  function setAiwordBatchStrategyMsg(msg, isErr) {
    if (!aiwordBatchStrategyMsg) return;
    aiwordBatchStrategyMsg.style.display = msg ? 'block' : 'none';
    aiwordBatchStrategyMsg.style.color = isErr ? 'var(--error)' : 'var(--text-muted)';
    aiwordBatchStrategyMsg.textContent = msg || '';
  }

  function _normBatchKey(v) {
    return String(v == null ? '' : v).trim();
  }

  function _ctxForFileId(fid) {
    return __aiwordHandoffCtxByFileId[String(fid)] || null;
  }

  function _ctxPhase(ctx) {
    if (!ctx || typeof ctx !== 'object') return '';
    var p = _normBatchKey(ctx.phase || ctx.task_type || ctx.belonging_module);
    return p || 'phase_unknown';
  }

  function _ctxAuthor(ctx) {
    if (!ctx || typeof ctx !== 'object') return '';
    return _normBatchKey(ctx.editor || ctx.writer);
  }

  function _ctxApprovals(ctx) {
    if (!ctx || typeof ctx !== 'object') return '|';
    return _normBatchKey(ctx.reviewer) + '|' + _normBatchKey(ctx.approver);
  }

  function _ctxDate(ctx) {
    if (!ctx || typeof ctx !== 'object') return '';
    return _normBatchKey(ctx.doc_date);
  }

  function _groupKeyByStrategy(fileId, strategy) {
    var ctx = _ctxForFileId(fileId) || {};
    var approvals = _ctxApprovals(ctx);
    var author = _ctxAuthor(ctx) || 'author_unknown';
    var phase = _ctxPhase(ctx);
    var dd = _ctxDate(ctx) || 'date_unknown';
    switch (strategy) {
      case 'approvals_phase':
        return approvals + '|' + phase;
      case 'approvals_only':
        return approvals;
      case 'author_only':
        return author;
      case 'approvals_phase_date':
      default:
        return approvals + '|' + phase + '|' + dd;
    }
  }

  function computeAiwordBatchGroups() {
    var strategy = aiwordBatchGroupStrategy ? String(aiwordBatchGroupStrategy.value || 'approvals_phase_date') : 'approvals_phase_date';
    var out = [];
    var idxByKey = {};
    savedFiles.forEach(function (f) {
      if (!f || !f.id) return;
      var fid = String(f.id);
      if (!__aiwordHandoffCtxByFileId[fid]) return;
      var gk = _groupKeyByStrategy(fid, strategy);
      if (idxByKey[gk] == null) {
        idxByKey[gk] = out.length;
        out.push({ key: gk, file_ids: [], names: [] });
      }
      var g = out[idxByKey[gk]];
      g.file_ids.push(fid);
      g.names.push(String(f.name || fid));
    });
    __aiwordBatchGroups = out;
    return out;
  }

  function renderAiwordBatchGroupsPreview() {
    if (!IS_FILE_SIGN_PAGE || !aiwordBatchStrategyCard) return;
    var groups = computeAiwordBatchGroups();
    if (!groups.length) {
      aiwordBatchStrategyCard.style.display = 'none';
      return;
    }
    aiwordBatchStrategyCard.style.display = 'block';
    if (aiwordBatchGroupList) {
      aiwordBatchGroupList.innerHTML = '';
      groups.forEach(function (g, idx) {
        var div = document.createElement('div');
        div.style.padding = '8px 10px';
        div.style.border = '1px solid var(--border)';
        div.style.borderRadius = '8px';
        div.style.marginBottom = '8px';
        div.innerHTML =
          '<strong>组 ' +
          (idx + 1) +
          '</strong> · ' +
          g.file_ids.length +
          ' 个文件<br><span class="hint">' +
          g.key.replace(/</g, '&lt;').replace(/>/g, '&gt;') +
          '</span>';
        if (Array.isArray(g.names) && g.names.length) {
          var ul = document.createElement('ul');
          ul.style.margin = '6px 0 0';
          ul.style.paddingLeft = '18px';
          g.names.forEach(function (nm) {
            var li = document.createElement('li');
            li.textContent = String(nm || '');
            ul.appendChild(li);
          });
          div.appendChild(ul);
        }
        aiwordBatchGroupList.appendChild(div);
      });
    }
    setAiwordBatchStrategyMsg('已按策略预览分组，共 ' + groups.length + ' 组。', false);
  }

  function exportAiwordBatchGroupsSnapshot() {
    var groups = computeAiwordBatchGroups();
    if (!groups.length) {
      setAiwordBatchStrategyMsg('当前无可导出的分组。', true);
      return;
    }
    var payload = {
      exported_at: new Date().toISOString(),
      strategy: aiwordBatchGroupStrategy ? String(aiwordBatchGroupStrategy.value || '') : '',
      groups: groups,
    };
    var blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json;charset=utf-8' });
    var a = document.createElement('a');
    var ts = new Date();
    var stamp = String(ts.getFullYear()) +
      String(ts.getMonth() + 1).padStart(2, '0') +
      String(ts.getDate()).padStart(2, '0') + '_' +
      String(ts.getHours()).padStart(2, '0') +
      String(ts.getMinutes()).padStart(2, '0') +
      String(ts.getSeconds()).padStart(2, '0');
    a.href = URL.createObjectURL(blob);
    a.download = 'aiword_batch_groups_' + stamp + '.json';
    a.click();
    setTimeout(function () {
      try {
        URL.revokeObjectURL(a.href);
      } catch (_) {}
    }, 1500);
    setAiwordBatchStrategyMsg('已导出分组快照。', false);
  }

  function applyAiwordBatchSelectionByGroups() {
    var groups = computeAiwordBatchGroups();
    if (!groups.length) {
      setAiwordBatchStrategyMsg('当前无可分组的 aiword 交接文件。', true);
      return;
    }
    var all = {};
    groups.forEach(function (g) {
      (g.file_ids || []).forEach(function (id) {
        all[String(id)] = true;
      });
    });
    document.querySelectorAll('.batch-pick').forEach(function (cb) {
      var id = String(cb.getAttribute('data-id') || '');
      cb.checked = !!all[id];
    });
    updateBatchUi();
    setAiwordBatchStrategyMsg('已按当前分组策略勾选 ' + Object.keys(all).length + ' 个文件。', false);
  }

  function _doBatchSignFromSubmit(opts) {
    opts = opts && typeof opts === 'object' ? opts : {};
    var ids = _checkedBatchFileIds();
    var fbWorkbench = opts.feedbackTarget === 'workbench';
    function setBatchFb(msg, isErr) {
      if (fbWorkbench) setBatchWorkbenchMsg(msg, isErr);
      else showErr(msg);
    }
    if (!ids.length) {
      setBatchFb('请先勾选要批量签名的文件', true);
      return;
    }
    if (!signersDbShare) {
      setBatchFb('批量签名需要启用 MySQL（MYSQL_HOST）', true);
      return;
    }
    var source = signSourceValue();
    if (source === 'library') {
      var pre = buildSignPreflightForFiles(ids);
      if (!pre.canProceed) {
        // 所有文件都没素材：明确给原因，并在工作台自动打开高级模式让用户能画布补签
        setBatchFb(
          '没有任何可签字的文件：\n  · ' + pre.blockers.join('\n  · ') +
            '\n\n可点击右上「高级模式」用画布手工补签，或先在「签名素材库」录入素材后重试。',
          'error'
        );
        if (fbWorkbench && !__batchWorkbenchAdvancedOpen && batchWorkbenchAdvancedCb) {
          batchWorkbenchAdvancedCb.checked = true;
          setBatchWorkbenchAdvanced(true);
        }
        return;
      }
      // 部分文件没素材：把它们从 ids 中剔除并提示用户；不再因为有 blocker 就整批拦截。
      if (pre.blockers.length) {
        var blockedSet = {};
        // 反推 blocker 对应的 fid：用 _fileDisplayNameById 反向匹配
        ids.forEach(function (fid) {
          var nm = _fileDisplayNameById(fid);
          pre.blockers.forEach(function (bk) {
            if (bk && nm && bk.indexOf(nm) === 0) blockedSet[String(fid)] = bk;
          });
        });
        var keepIds = ids.filter(function (fid) { return !blockedSet[String(fid)]; });
        if (!keepIds.length) {
          setBatchFb(
            '所选文件均无可用素材：\n  · ' + pre.blockers.join('\n  · '),
            'error'
          );
          return;
        }
        // 把被剔除的文件状态标到 row 上
        Object.keys(blockedSet).forEach(function (fid) {
          var row = __batchWorkbenchRows[fid];
          if (row && (!row.status || row.status === '识别有误' || row.status === '待匹配' || row.status === '—')) {
            row.status = '无可用素材';
            row.rolesLabel = blockedSet[fid];
          }
        });
        setBatchFb(
          '以下文件无任何角色素材，已跳过（共 ' + Object.keys(blockedSet).length + ' 个）；其余 ' +
            keepIds.length + ' 个将继续签字：\n  · ' + pre.blockers.slice(0, 8).join('\n  · ') +
            (pre.blockers.length > 8 ? '\n  …' : ''),
          'warn'
        );
        ids = keepIds;
        try { renderBatchWorkbenchTable(); } catch (_) {}
      }
      if (
        (pre.partial.length || pre.compositeWarn.length) &&
        !confirmPartialSignProceed(pre)
      ) {
        setBatchFb('已取消生成', false);
        return;
      }
      if (fbWorkbench && (pre.partial.length || pre.compositeWarn.length)) {
        setBatchWorkbenchMsg('将为已匹配字段生成；不完整项将自动跳过，可在结果里看到具体明细', 'warn');
      }
    }
    var libRestrict = libraryRolesRestrictedToChecks();
    var rolesForBatch = null;
    if (source === 'library' && !libRestrict) {
      rolesForBatch = null;
    } else {
      rolesForBatch = selectedRoleIds();
      if (!rolesForBatch.length) {
        showErr('请至少勾选一个角色（用于批量签）');
        return;
      }
    }
    var payload = {
      file_ids: ids,
      source: source,
      apply_person: opts.apply_person !== false,
      apply_date: opts.apply_date !== false,
    };
    if (rolesForBatch != null) {
      payload.roles = rolesForBatch;
    }
    if (source === 'canvas') {
      resizeCanvasesForRoles(rolesForBatch);
      payload.sig_map = {};
      payload.date_map = {};
      rolesForBatch.forEach(function (rid3) {
        var sigC3 = document.getElementById('sig_' + rid3);
        var dateC3 = document.getElementById('date_' + rid3);
        if (sigC3 && !isCanvasBlank(sigC3)) {
          payload.sig_map[rid3] = _normalizedPngDataUrl(sigC3, 'sig');
        }
        if (dateC3 && !isCanvasBlank(dateC3)) {
          payload.date_map[rid3] = _normalizedPngDataUrl(dateC3, 'date');
        }
      });
    }
    showErr('');
    if (!opts.skipPageProgress && !isBatchWorkbenchMode()) {
      beginPageProgress('批量生成签字文档…', { done: 0, total: ids.length });
    }
    if (submitBtn) {
      submitBtn.disabled = true;
      submitBtn.innerHTML = '<span class="spinner"></span> 批量处理中…';
    }
    if (fbWorkbench && batchWorkbenchSignBtn) {
      batchWorkbenchSignBtn.disabled = true;
      batchWorkbenchSignBtn.dataset._origText = batchWorkbenchSignBtn.dataset._origText || batchWorkbenchSignBtn.textContent || '';
      batchWorkbenchSignBtn.textContent = '签字中…';
    }
    var persist0 = source === 'library' ? bulkPersistRoleMapsForBatch(ids) : Promise.resolve([]);
    var total = ids.length;
    var t0 = Date.now();
    var batchId = _randomSignBatchIdHex32();
    var allResults = [];
    // 工作台模式：把每文件签字状态回写到 row.status，让用户立刻看到结果
    function markRow(fid, status, label) {
      if (!fbWorkbench || !isBatchWorkbenchMode()) return;
      var row = __batchWorkbenchRows[String(fid)];
      if (!row) return;
      row.status = status;
      if (label) row.rolesLabel = label;
      try { renderBatchWorkbenchTable(); } catch (_) {}
    }

    function signBatchProgressLine(done, phaseSuffix) {
      var elapsed = (Date.now() - t0) / 1000;
      var pct = total ? Math.min(100, Math.round((done / total) * 100)) : 0;
      var tail = '';
      if (done > 0 && done < total) {
        var avg = elapsed / done;
        tail =
          ' · 预计剩余 ' +
          _formatSignBatchEtaSeconds(avg * (total - done)) +
          '（近均 ' +
          (Math.round(avg * 10) / 10) +
          ' 秒/文件）';
      } else if (done >= total && total) {
        tail = ' · 本批总耗时 ' + _formatSignBatchEtaSeconds(elapsed);
      }
      var head = '批量签字 进度 ' + done + '/' + total + '（' + pct + '%）';
      var line = head + (phaseSuffix ? ' ' + phaseSuffix : '') + ' · 已用 ' + _formatSignBatchEtaSeconds(elapsed) + tail;
      showBatchResult(line, false);
      if (fbWorkbench) setBatchFb(line, 'ok');
      if (fbWorkbench) setBatchWorkbenchProgress(done, total, line, true);
    }

    return persist0
      .then(function (persistReport) {
        // role-map 写库阶段：失败项不阻塞整批，但要让用户看到具体哪个文件 PUT 失败
        if (Array.isArray(persistReport)) {
          var persistFails = persistReport.filter(function (r) { return r && !r.ok; });
          if (persistFails.length) {
            persistFails.forEach(function (r) {
              markRow(r.fid, '保存映射失败', r.error || 'role-map 保存失败');
            });
            var failNames = persistFails.map(function (r) {
              return (_fileDisplayNameById(r.fid) || r.fid) + '：' + (r.error || '保存失败');
            });
            setBatchFb(
              '部分文件角色映射保存失败，但仍会尝试用服务端旧映射继续签字：\n  · ' +
                failNames.slice(0, 8).join('\n  · ') +
                (failNames.length > 8 ? '\n  …其余 ' + (failNames.length - 8) + ' 项' : ''),
              'warn'
            );
          }
        }
        signBatchProgressLine(0, '· 正在提交第 1 个文件…');
        var chain2 = Promise.resolve();
        ids.forEach(function (fid, idx) {
          chain2 = chain2.then(function () {
            var rec = savedFiles.find(function (x) {
              return x && String(x.id) === String(fid);
            });
            var nm = (rec && (rec.name || rec.id)) ? String(rec.name || rec.id) : String(fid);
            markRow(fid, '签字中…', '正在生成签字版文档');
            signBatchProgressLine(idx, '· 正在处理「' + nm + '」…');
            var onePayload = Object.assign({}, payload, { file_ids: [fid], batch_id: batchId });
            if (opts.workbenchMode && isBatchWorkbenchMode()) {
              onePayload.source = 'library';
              var wbOv = collectWorkbenchCanvasOverridesForFile(fid);
              if (wbOv.sig_map && Object.keys(wbOv.sig_map).length) {
                onePayload.sig_map = wbOv.sig_map;
              }
              if (wbOv.date_map && Object.keys(wbOv.date_map).length) {
                onePayload.date_map = wbOv.date_map;
              }
            }
            return fetchJson(apiUrl('/api/sign/batch'), {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify(onePayload),
              timeoutMs: 600000,
            }).then(function (r) {
              var j = r.data || {};
              if (j.batch_id) batchId = String(j.batch_id);
              if (!j.ok) {
                // 整个请求级失败（如 503 server_busy）：把这个文件标记为失败，
                // 但继续下一个，不再 throw 中断整批。
                var errMsg = j.error || '接口未返回 ok';
                allResults.push({ file_id: fid, ok: false, name: nm, error: errMsg });
                markRow(fid, '签字失败', errMsg);
                signBatchProgressLine(idx + 1, '· 「' + nm + '」失败：' + errMsg);
                return;
              }
              var chunk = j.results || [];
              for (var ci = 0; ci < chunk.length; ci++) {
                var it = chunk[ci];
                if (!it) continue;
                if (!it.name) it.name = nm;
                allResults.push(it);
                if (String(it.file_id) === String(fid)) {
                  if (it.ok) {
                    var okLbl = '已签字' + (it.applied_n ? '（' + it.applied_n + ' 项）' : '');
                    markRow(fid, '已签字', okLbl);
                  } else {
                    markRow(fid, '签字失败', it.error || '签字失败');
                  }
                }
              }
              signBatchProgressLine(idx + 1, '· 「' + nm + '」已返回');
            }).catch(function (err) {
              var msg = (err && err.message) || String(err || '');
              allResults.push({ file_id: fid, ok: false, name: nm, error: msg });
              markRow(fid, '签字失败', msg);
              signBatchProgressLine(idx + 1, '· 「' + nm + '」失败：' + msg);
              // 关键：单文件失败时**不再 throw**，让链继续处理下一个文件
            });
          });
        });
        return chain2;
      })
      .then(function () {
        var res = allResults;
        var okn = res.filter(function (x) { return x && x.ok; }).length;
        var failn = res.length - okn;
        var lines = [];
        lines.push(
          '批量完成：成功 ' + okn + ' / ' + res.length + '（批次 ' + (batchId ? batchId.slice(0, 8) : '') + '…）。'
        );
        res.forEach(function (it) {
          if (!it) return;
          lines.push(formatBatchSignItemResult(it));
        });
        showBatchResult(lines.join('\n\n'), false);
        if (fbWorkbench) {
          var head2 = '批量签字完成：成功 ' + okn + ' / ' + res.length + '。';
          if (failn) {
            head2 += '\n失败 ' + failn + ' 个文件，详见每行「状态」列；可在「已签文档列表」单独下载成功项。';
            setBatchFb(head2, 'warn');
          } else {
            setBatchFb(head2 + '\n可在下方「已签文档列表」下载。', 'ok');
          }
        }
        refreshSignedList({ skipPageProgress: true });
      })
      .catch(function (e) {
        // 进入这里说明 persist0 或 chain2 之外的代码抛错（如 setup 阶段）。
        var msg = (e && e.message) || String(e || '');
        setBatchFb('批量签字流程出错：' + msg + '\n已完成 ' + allResults.length + '/' + total +
          '，可重新点击「生成签字版文档」重试。', 'error');
        signBatchProgressLine(allResults.length, '· 已中断（已完成 ' + allResults.length + '/' + total + '）');
      })
      .then(function () {
        if (submitBtn) {
          submitBtn.disabled = false;
        }
        if (fbWorkbench && batchWorkbenchSignBtn) {
          batchWorkbenchSignBtn.disabled = false;
          batchWorkbenchSignBtn.textContent = batchWorkbenchSignBtn.dataset._origText || '生成签字版文档';
        }
        if (fbWorkbench) setBatchWorkbenchProgress(total, total, '批量签字完成 ' + total + '/' + total + '（100%）', true);
        updateSubmitState();
        updateBatchUi();
      })
      .finally(function () {
        if (!opts.skipPageProgress) endPageProgress();
      });
  }

  function buildUI() {
    // 两个页面复用同一脚本：按页面类型只初始化各自需要的 UI
    if (!IS_FILE_SIGN_PAGE && !IS_MATERIALS_PAGE) return;

    if (IS_FILE_SIGN_PAGE) ROLES.forEach(function (r) {
      var row = document.createElement('div');
      row.className = 'role-row';
      var chk = document.createElement('input');
      chk.type = 'checkbox';
      chk.id = 'chk_' + r.id;
      chk.setAttribute('data-id', r.id);
      var lbl = document.createElement('label');
      lbl.htmlFor = chk.id;
      lbl.style.margin = '0';
      lbl.style.cursor = 'pointer';
      lbl.textContent = r.label;
      row.appendChild(chk);
      row.appendChild(lbl);
      roleChecks.appendChild(row);

      var panel = document.createElement('div');
      panel.className = 'role-panel';
      panel.id = 'panel_' + r.id;

      function addCanvasBlock(titlePrefix, clearLabel, canvasId) {
        var wrap = document.createElement('div');
        wrap.className = 'canvas-wrap';
        var lab = document.createElement('label');
        lab.textContent = titlePrefix + '（' + r.label + '）';
        var cvs = document.createElement('canvas');
        cvs.className = 'sign-pad';
        cvs.id = canvasId;
        var br = document.createElement('div');
        br.className = 'btn-row';
        var btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'btn btn-secondary';
        btn.setAttribute('data-clear', canvasId);
        btn.textContent = clearLabel;
        br.appendChild(btn);
        wrap.appendChild(lab);
        wrap.appendChild(cvs);
        wrap.appendChild(br);

        // 画布旁入库并绑定（按角色、按 kind）
        var kind = /^date_/.test(canvasId) ? 'date' : 'sig';
        var rid = r.id;
        var state = _ensureRoleSaveSignerState(rid)[kind];
        var loc = _roleFinalLocale(rid);

        var filter = document.createElement('input');
        filter.type = 'search';
        filter.placeholder = '筛选签署人（保存到谁）…';
        filter.value = state.q || '';
        filter.style.width = '100%';
        filter.style.boxSizing = 'border-box';
        filter.style.padding = '6px';
        filter.style.marginTop = '6px';
        filter.style.border = '1px solid var(--border)';
        filter.style.borderRadius = '8px';

        var sel = document.createElement('select');
        var saveWrap = document.createElement('div');
        saveWrap.className = 'role-save-signer';
        saveWrap.setAttribute('data-rid', rid);
        saveWrap.setAttribute('data-kind', kind);
        sel.style.width = '100%';
        sel.style.maxWidth = '100%';
        sel.style.padding = '6px';
        sel.style.marginTop = '6px';
        sel.className = 'role-save-signer-select';
        fillSignerSelect(sel, state.id || '', state.q || '');
        filter.addEventListener('input', function () {
          state.q = filter.value || '';
          var prev = sel.value || state.id || '';
          fillSignerSelect(sel, prev, state.q);
        });
        sel.addEventListener('change', function () {
          state.id = sel.value || '';
        });
        sel.addEventListener('focus', function () {
          // 防止页面初次加载时 signersList 为空导致下拉为空：获得焦点时即时回填一次
          fillSignerSelect(sel, sel.value || state.id || '', state.q || (filter.value || ''), _roleFinalLocale(rid));
        });
        filter.classList.add('role-save-signer-filter');

        var saveBtn = document.createElement('button');
        saveBtn.type = 'button';
        saveBtn.className = 'btn btn-secondary';
        saveBtn.style.marginTop = '6px';
        saveBtn.textContent = (kind === 'date' ? '入库日期并绑定本角色' : '入库签名并绑定本角色');
        var panelSaveFb = document.createElement('p');
        panelSaveFb.className = 'btn-inline-feedback';
        panelSaveFb.style.display = 'none';
        panelSaveFb.style.marginTop = '6px';
        panelSaveFb.setAttribute('role', 'status');
        saveBtn.addEventListener('click', function () {
          var signerForSave = sel.value || state.id || '';
          if (!signerForSave) {
            setPanelSaveFeedback(panelSaveFb, '请先选择要保存到的签署人（可在上方添加新签署人）', true);
            return;
          }
          var finalLoc = _roleFinalLocale(rid);
          // 覆盖确认：该签署人该语言该类别已存在
          try {
            var s0 = signersList.find(function (x) { return x.id === signerForSave; });
            if (s0 && _signerHasKindInLocale(s0, kind, finalLoc)) {
              var nm = s0.name || s0.id;
              if (!window.confirm('“' + nm + '” 已有 ' + (finalLoc === 'en' ? '英文' : '中文') + ' 的' + (kind === 'date' ? '日期' : '签名') + '图片，是否覆盖？')) {
                return;
              }
            }
          } catch (_) {}

          setRoleChecked(rid, true);
          requestAnimationFrame(function () {
            requestAnimationFrame(function () {
              resizeCanvasesForRoles([rid]);
              var c = document.getElementById(canvasId);
              if (isCanvasBlank(c)) {
                setPanelSaveFeedback(
                  panelSaveFb,
                  '请先在「' + roleLabel(rid) + '」' + (kind === 'date' ? '日期' : '签名') + '画布上手写',
                  true
                );
                return;
              }
              var fd = new FormData();
              fd.append(kind, _normalizedPngDataUrl(c, kind));
              fd.append('locale', finalLoc);
              withButtonBusy(saveBtn, '保存中…', function () {
                return fetchJson(apiUrl('/api/sign/signers/' + signerForSave + '/strokes'), {
                  method: 'PUT',
                  body: fd,
                }).then(function (rr) {
                  var jj = rr.data || {};
                  if (!jj.ok) {
                    setPanelSaveFeedback(panelSaveFb, jj.error || '保存失败', true);
                    return;
                  }
                  setPanelSaveFeedback(panelSaveFb, '', false);
                  var newId = (kind === 'date') ? jj.date_item_id : jj.sig_item_id;
                  var fidPanel = selectedFileId;
                  if (newId && fidPanel) {
                    var m = Object.assign({}, currentRoleMap);
                    var p = m[rid] && typeof m[rid] === 'object' ? Object.assign({}, m[rid]) : {};
                    if (kind === 'date') p.date = newId;
                    else p.sig = newId;
                    m[rid] = p;
                    currentRoleMap = m;
                    cachePatchCurrentRoleMap(fidPanel, currentRoleMap);
                    return fetchJson(apiUrl('/api/sign/files/' + fidPanel + '/role-map'), {
                      method: 'PUT',
                      headers: { 'Content-Type': 'application/json' },
                      body: JSON.stringify({ map: m }),
                    }).then(function (r2) {
                      var j2 = r2.data;
                      if (j2 && j2.ok) {
                        if (String(selectedFileId) === String(fidPanel)) {
                          currentRoleMap = j2.map || m;
                        }
                        cachePatchCurrentRoleMap(fidPanel, j2.map || m);
                      }
                      setPanelSaveFeedback(
                        panelSaveFb,
                        (kind === 'date' ? '日期' : '签名') + '已入库并已绑定本角色。',
                        false
                      );
                      return refreshSigners();
                    }).then(function () {
                      renderNeedSignTable();
                    });
                  }
                  setPanelSaveFeedback(
                    panelSaveFb,
                    (kind === 'date' ? '日期' : '签名') + '已保存到所选签署人。',
                    false
                  );
                  return refreshSigners();
                });
              }).catch(function (e) {
                setPanelSaveFeedback(panelSaveFb, e.message || String(e), true);
              });
            });
          });
        });

        saveWrap.appendChild(filter);
        saveWrap.appendChild(sel);
        saveWrap.appendChild(saveBtn);
        saveWrap.appendChild(panelSaveFb);
        wrap.appendChild(saveWrap);
        panel.appendChild(wrap);
      }
      addCanvasBlock('签名', '清除签名', 'sig_' + r.id);
      addCanvasBlock('日期', '清除日期', 'date_' + r.id);
      rolePanels.appendChild(panel);

      chk.addEventListener('change', function () {
        var vis = chk.checked;
        panel.classList.toggle('visible', vis);
        if (vis) {
          requestAnimationFrame(function () {
            requestAnimationFrame(function () {
              resizeCanvasesForRoles([r.id]);
              try {
                panel.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'nearest' });
              } catch (e) {}
            });
          });
        }
        if (IS_FILE_SIGN_PAGE && selectedFileId) saveCurrentFileUiToCache(selectedFileId);
        renderNeedSignTable();
        updateSubmitState();
      });
    });

    if (IS_FILE_SIGN_PAGE) rolePanels.querySelectorAll('[data-clear]').forEach(function (btn) {
      btn.addEventListener('click', function () {
        var id = btn.getAttribute('data-clear');
        if (canvases[id]) canvases[id].clear();
      });
    });

    if (IS_FILE_SIGN_PAGE) ROLES.forEach(function (r) {
      canvases['sig_' + r.id] = setupCanvas(document.getElementById('sig_' + r.id), {
        lineWidth: 4.35,
        shadowBlur: 0.7,
      });
      canvases['date_' + r.id] = setupCanvas(document.getElementById('date_' + r.id), {
        lineWidth: 3.55,
        shadowBlur: 0.55,
      });
    });

    if (IS_MATERIALS_PAGE) canvases['lib_sig_canvas'] = setupCanvas(document.getElementById('lib_sig_canvas'), {
      lineWidth: 4.35,
      shadowBlur: 0.7,
    });
    if (IS_MATERIALS_PAGE) canvases['lib_date_canvas'] = setupCanvas(document.getElementById('lib_date_canvas'), {
      lineWidth: 3.55,
      shadowBlur: 0.55,
    });

    if (IS_MATERIALS_PAGE) (function initPieceDateUi() {
      var MONTHS_EN = [
        'January',
        'February',
        'March',
        'April',
        'May',
        'June',
        'July',
        'August',
        'September',
        'October',
        'November',
        'December',
      ];
      var MONTH_ABBREV = [
        'Jan',
        'Feb',
        'Mar',
        'Apr',
        'May',
        'Jun',
        'Jul',
        'Aug',
        'Sep',
        'Oct',
        'Nov',
        'Dec',
      ];

      var PIECE_ORDER = [];
      var pi;
      for (pi = 0; pi <= 9; pi++) {
        PIECE_ORDER.push({ k: 'pd' + pi, lab: '数字 ' + pi });
      }
      for (pi = 0; pi < 12; pi++) {
        var n2 = pi + 1;
        var kva = 'pma' + (n2 < 10 ? '0' : '') + n2;
        PIECE_ORDER.push({
          k: kva,
          lab: '月份 ' + n2 + '（' + MONTH_ABBREV[pi] + '）',
        });
      }
      PIECE_ORDER.push({ k: 'pdot', lab: '句点 .' });

      function appendPieceChecks(containerEl, metas) {
        if (!containerEl) return;
        containerEl.innerHTML = '';
        (metas || []).forEach(function (meta) {
          var lab = document.createElement('label');
          lab.style.display = 'inline-flex';
          lab.style.alignItems = 'center';
          lab.style.gap = '4px';
          lab.style.fontSize = '0.82rem';
          lab.style.cursor = 'pointer';
          var cb = document.createElement('input');
          cb.type = 'checkbox';
          cb.setAttribute('data-kind', meta.k);
          cb.title = meta.lab;
          var sp = document.createElement('span');
          sp.textContent = meta.lab;
          lab.appendChild(cb);
          lab.appendChild(sp);
          containerEl.appendChild(lab);
        });
      }

      var digitMetas = PIECE_ORDER.filter(function (x) {
        return /^pd[0-9]$/.test(x.k) || x.k === 'pdot';
      });
      var monthMetas = PIECE_ORDER.filter(function (x) {
        return /^pma(0[1-9]|1[0-2])$/.test(x.k);
      });

      appendPieceChecks(pieceDigitChecks, digitMetas);
      appendPieceChecks(pieceMonthChecks, monthMetas);

      var batchOrder = [];
      var batchStep = 0;
      var batchQueue = [];
      var batchActive = false;
      var PIECE_BATCH_DRAFT_KEY = 'sign_piece_batch_draft_v1';
      var __pieceDraftRestored = false;
      var pieceBatchUploadAc = null;
      var pieceBatchUploadUserCancel = false;
      var PIECE_BATCH_CANCEL_LABEL_IDLE = '取消批量';
      var PIECE_BATCH_CANCEL_LABEL_BUSY = '取消上传';

      function _isPieceBatchUploading() {
        return !!(pieceBatchUploadBtn && pieceBatchUploadBtn.getAttribute('aria-busy') === 'true');
      }

      function _setPieceBatchCancelLabel(uploading) {
        if (!pieceBatchCancelBtn) return;
        pieceBatchCancelBtn.textContent = uploading
          ? PIECE_BATCH_CANCEL_LABEL_BUSY
          : PIECE_BATCH_CANCEL_LABEL_IDLE;
        pieceBatchCancelBtn.title = uploading
          ? '中止当前网络上传，队列与草稿会保留，可稍后点「上传队列中全部」重试'
          : '放弃本次批量录入：清空未上传队列与本地草稿（已入库的元件不受影响）';
      }

      function _beginPieceBatchUploadAbort() {
        pieceBatchUploadUserCancel = false;
        if (pieceBatchUploadAc) {
          try {
            pieceBatchUploadAc.abort();
          } catch (_) {}
        }
        pieceBatchUploadAc =
          typeof AbortController !== 'undefined' ? new AbortController() : null;
        _setPieceBatchCancelLabel(true);
      }

      function _endPieceBatchUploadAbort() {
        pieceBatchUploadAc = null;
        pieceBatchUploadUserCancel = false;
        _setPieceBatchCancelLabel(false);
      }

      function _pieceDraftSnapshot(reason) {
        return {
          v: 1,
          sid: getLibActiveSignerId(),
          signer_name: (function () {
            var id = getLibActiveSignerId();
            var s = id ? _findSignerById(id) : null;
            return s ? String(s.name || s.id || id) : '';
          })(),
          batchOrder: (batchOrder || []).slice(),
          batchStep: Number(batchStep || 0),
          batchQueue: (batchQueue || []).slice(),
          batchActive: !!batchActive,
          reason: reason || '',
          ts: Date.now(),
        };
      }

      function savePieceBatchDraft(reason) {
        try {
          if (!window.localStorage) return;
          var snap = _pieceDraftSnapshot(reason);
          if (!(snap.batchQueue && snap.batchQueue.length)) {
            return;
          }
          window.localStorage.setItem(PIECE_BATCH_DRAFT_KEY, JSON.stringify(snap));
        } catch (_) {}
      }

      function clearPieceBatchDraft() {
        try {
          if (!window.localStorage) return;
          window.localStorage.removeItem(PIECE_BATCH_DRAFT_KEY);
        } catch (_) {}
      }

      function tryRestorePieceBatchDraft() {
        if (__pieceDraftRestored) return;
        __pieceDraftRestored = true;
        try {
          if (!window.localStorage) return;
          var raw = window.localStorage.getItem(PIECE_BATCH_DRAFT_KEY);
          if (!raw) return;
          var d = JSON.parse(raw || '{}');
          if (!d || !Array.isArray(d.batchQueue) || !d.batchQueue.length) return;
          var ageSec = Math.max(0, Math.floor((Date.now() - Number(d.ts || 0)) / 1000));
          var msg =
            '检测到一份元件上传草稿（' +
            d.batchQueue.length +
            ' 项，约 ' +
            ageSec +
            ' 秒前）' +
            (d.signer_name ? '，签署人：' + d.signer_name : '') +
            '。是否恢复并重试上传？';
          if (!window.confirm(msg)) return;
          batchOrder = Array.isArray(d.batchOrder) ? d.batchOrder.slice() : [];
          batchQueue = d.batchQueue.slice();
          batchStep = Number(d.batchStep || batchQueue.length);
          batchActive = false;
          if (d.sid) {
            setLibActiveSignerId(d.sid);
          }
          pieceBatchNextBtn.disabled = true;
          pieceBatchUploadBtn.disabled = false;
          pieceBatchStartBtn.disabled = false;
          setPieceBatchFeedback(
            '已恢复上传草稿：' + batchQueue.length + ' 项。可直接点「上传队列中全部」重试。',
            false
          );
        } catch (_) {}
      }
      try { window.__tryRestorePieceDraft = tryRestorePieceBatchDraft; } catch (_) {}

      function pieceMetaLabel(kind) {
        var f = PIECE_ORDER.find(function (x) {
          return x.k === kind;
        });
        var lab = f ? f.lab : kind;
        if (kind === 'pd0') return lab + '（注意：这是数字 0，不是句点）';
        if (kind === 'pdot') return lab + '（注意：这是句点 .，不是数字 0）';
        return lab;
      }

      function resetPieceBatchUi() {
        batchOrder = [];
        batchStep = 0;
        batchQueue = [];
        batchActive = false;
        clearPieceBatchDraft();
        pieceBatchNextBtn.disabled = true;
        pieceBatchUploadBtn.disabled = true;
        pieceBatchStartBtn.disabled = false;
        setPieceBatchFeedback('', false);
        pieceDigitChecks.querySelectorAll('input[type=checkbox]').forEach(function (x) {
          x.disabled = false;
          x.checked = false;
        });
        pieceMonthChecks.querySelectorAll('input[type=checkbox]').forEach(function (x) {
          x.disabled = false;
          x.checked = false;
        });
      }

      canvases['piece_canvas'] = setupCanvas(pieceCanvasEl, {
        lineWidth: 3.55,
        shadowBlur: 0.55,
      });
      pieceClearBtn.addEventListener('click', function () {
        if (canvases['piece_canvas'] && canvases['piece_canvas'].clear) canvases['piece_canvas'].clear();
      });

      pieceBatchStartBtn.addEventListener('click', function () {
        if (!signersDbShare) {
          setPieceBatchFeedback('笔迹元件需启用 MySQL', true);
          return;
        }
        if (!getLibActiveSignerId()) {
          setPieceBatchFeedback(
            '请先在上方选择「当前签署人」（当前：' +
              (currentSignerName.textContent || '—') +
              '）',
            true
          );
          return;
        }
        batchOrder = [];
        [pieceDigitChecks, pieceMonthChecks].forEach(function (wrap) {
          if (!wrap) return;
          wrap.querySelectorAll('input[type=checkbox]').forEach(function (cb) {
            if (cb && cb.checked) {
              var k = cb.getAttribute('data-kind') || '';
              if (k) batchOrder.push(k);
            }
          });
        });
        if (!batchOrder.length) {
          setPieceBatchFeedback('请至少勾选一个元件', true);
          return;
        }
        batchQueue = [];
        batchStep = 0;
        batchActive = true;
        savePieceBatchDraft('started');
        pieceBatchNextBtn.disabled = false;
        pieceBatchUploadBtn.disabled = true;
        pieceBatchStartBtn.disabled = true;
        pieceDigitChecks.querySelectorAll('input[type=checkbox]').forEach(function (x) {
          x.disabled = true;
        });
        pieceMonthChecks.querySelectorAll('input[type=checkbox]').forEach(function (x) {
          x.disabled = true;
        });
        clearFileRegionErr();
        setPieceBatchProgress(
          '第 1/' + batchOrder.length + ' 项：' + pieceMetaLabel(batchOrder[0]) + ' — 请在下方手写后点「下一项」。'
        );
      });

      pieceBatchNextBtn.addEventListener('click', function () {
        if (!batchActive || batchStep >= batchOrder.length) return;
        var kind = batchOrder[batchStep];
        if (isCanvasBlank(pieceCanvasEl)) {
          setPieceBatchFeedback('请先在手写区书写「' + pieceMetaLabel(kind) + '」', true);
          return;
        }
        batchQueue.push({
          piece_kind: kind,
          png: _normalizedPngDataUrl(pieceCanvasEl, 'date'),
        });
        savePieceBatchDraft('collecting');
        if (canvases['piece_canvas'] && canvases['piece_canvas'].clear) canvases['piece_canvas'].clear();
        batchStep += 1;
        clearFileRegionErr();
        if (batchStep >= batchOrder.length) {
          batchActive = false;
          pieceBatchNextBtn.disabled = true;
          pieceBatchUploadBtn.disabled = false;
          savePieceBatchDraft('ready_to_upload');
          setPieceBatchProgress(
            '已将 ' +
              batchQueue.length +
              ' 项加入上传队列，请点击「上传队列中全部」。'
          );
        } else {
          setPieceBatchProgress(
            '第 ' +
              (batchStep + 1) +
              '/' +
              batchOrder.length +
              ' 项：' +
              pieceMetaLabel(batchOrder[batchStep]) +
              ' — 请书写后点「下一项」。'
          );
        }
      });

      pieceBatchUploadBtn.addEventListener('click', function () {
        if (!signersDbShare) {
          setPieceBatchFeedback('笔迹元件需启用 MySQL（db_share）', true);
          return;
        }
        var sid = String(libActiveSignerId || '').trim();
        if (!sid) sid = getLibActiveSignerId();
        if (!sid) {
          setPieceBatchFeedback(
            '请先在上方选择「当前签署人」（当前：' +
              (currentSignerName.textContent || '—') +
              '）',
            true
          );
          return;
        }
        if (signersList.length && !_findSignerById(sid)) {
          setPieceBatchFeedback('所选签署人已不存在，请重新选择', true);
          return;
        }
        var sidFrozen = sid;
        if (libSignerSelect && libSignerSelect.value && libSignerSelect.value !== sidFrozen) {
          setLibActiveSignerId(sidFrozen);
        }
        if (!batchQueue.length) {
          setPieceBatchFeedback('没有可上传的队列（请先完成批量录入）', true);
          return;
        }
        function _pieceBatchBusyLabel(text) {
          if (!pieceBatchUploadBtn || pieceBatchUploadBtn.getAttribute('aria-busy') !== 'true') {
            return;
          }
          pieceBatchUploadBtn.innerHTML =
            '<span class="spinner" aria-hidden="true"></span> ' + (text || '上传中…');
        }

        function _pieceBatchItemDesc(item) {
          var k = item && item.piece_kind ? String(item.piece_kind) : '';
          if (!k) return '未知元件';
          return pieceMetaLabel(k) + '（' + k + '）';
        }

        function _reportPieceUploadProgress(cur, total, item, phaseNote, counts) {
          counts = counts || {};
          var desc = _pieceBatchItemDesc(item);
          var pct = total > 0 ? Math.round((cur / total) * 100) : 0;
          var okn = counts.ok != null ? counts.ok : 0;
          var failn = counts.fail != null ? counts.fail : 0;
          var line =
            '正在上传 ' +
            cur +
            '/' +
            total +
            '（' +
            pct +
            '%）：' +
            desc +
            (phaseNote ? ' — ' + phaseNote : '') +
            (cur > 1 || okn || failn ? '；累计成功 ' + okn + '，失败 ' + failn : '');
          setPieceBatchProgress(line);
          setPieceBatchFeedback(line, false);
          _pieceBatchBusyLabel(cur + '/' + total);
          try {
            if (pieceBatchStatus) {
              pieceBatchStatus.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
            }
          } catch (_) {}
        }

        function uploadPiecesSequential(overwrite) {
          var sidPut = String(libActiveSignerId || '').trim() || sidFrozen;
          var total = batchQueue.length;
          var allResults = [];
          var okSoFar = 0;
          var failSoFar = 0;
          _beginPieceBatchUploadAbort();
          _reportPieceUploadProgress(0, total, batchQueue[0], '准备中…', { ok: 0, fail: 0 });

          function uploadAtIndex(i) {
            if (pieceBatchUploadUserCancel) {
              return Promise.reject(_uploadCancelledError());
            }
            if (i >= total) {
              return Promise.resolve({
                res: { ok: true },
                data: { ok: true, results: allResults },
              });
            }
            var item = batchQueue[i];
            var cur = i + 1;
            _reportPieceUploadProgress(cur, total, item, '提交中…', {
              ok: okSoFar,
              fail: failSoFar,
            });
            savePieceBatchDraft('uploading_' + cur + '_' + (item.piece_kind || ''));

            var fetchOpts = {
              method: 'PUT',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ items: [item], overwrite: !!overwrite }),
              timeoutMs: PIECE_BATCH_ITEM_TIMEOUT_MS,
            };
            if (pieceBatchUploadAc) {
              fetchOpts.signal = pieceBatchUploadAc.signal;
            }
            return fetchJsonWithRetry(
              apiUrl('/api/sign/signers/' + sidPut + '/stroke-pieces'),
              fetchOpts,
              {
                maxTry: 3,
                delayMs: 1000,
                onRetry: function (nextTry, maxTry2) {
                  savePieceBatchDraft('upload_retry');
                  _reportPieceUploadProgress(cur, total, item, '网络重试 ' + nextTry + '/' + maxTry2, {
                    ok: okSoFar,
                    fail: failSoFar,
                  });
                },
              }
            ).then(function (res) {
              var jj = res.data || {};
              if (!jj.ok) {
                var busyHint =
                  jj.error_code === 'server_busy' || res.res.status === 503
                    ? '（另一客户端可能正在批量签名，请稍后再传本条）'
                    : '';
                throw new Error(
                  (jj.error ||
                    ('上传失败' + (res.res && !res.res.ok ? '（HTTP ' + res.res.status + '）' : ''))) +
                    busyHint
                );
              }
              var row =
                (jj.results && jj.results[0]) ||
                { ok: false, piece_kind: item.piece_kind, error: '服务器未返回本条结果' };
              allResults.push(row);
              if (row.ok) {
                okSoFar += 1;
                _reportPieceUploadProgress(cur, total, item, '本条已保存', {
                  ok: okSoFar,
                  fail: failSoFar,
                });
              } else {
                failSoFar += 1;
                _reportPieceUploadProgress(
                  cur,
                  total,
                  item,
                  '本条未保存：' + (row.error || row.error_code || '失败'),
                  { ok: okSoFar, fail: failSoFar }
                );
              }
              return uploadAtIndex(i + 1);
            });
          }

          return uploadAtIndex(0);
        }

        function _refreshSignersAfterPieceUpload() {
          scheduleRefreshSigners(1200);
        }

        function handleUploadResult(res) {
          var jj = res.data || {};
          if (!jj.ok) {
            setPieceBatchFeedback(
              jj.error ||
                ('批量保存失败' +
                  (res.res && !res.res.ok ? '（HTTP ' + res.res.status + '）' : '')),
              true
            );
            savePieceBatchDraft('upload_failed');
            _refreshSignersAfterPieceUpload();
            return Promise.resolve();
          }
          var okn = 0;
          var failn = 0;
          var firstErrs = [];
          var hasExists = false;
          (jj.results || []).forEach(function (r) {
            if (r && r.ok) okn += 1;
            else {
              failn += 1;
              if (r && (r.error_code === 'exists' || /已存在|同名元件/.test(String(r.error || '')))) {
                hasExists = true;
              }
              if (firstErrs.length < 3 && r && (r.error || r.piece_kind)) {
                firstErrs.push(
                  (r.piece_kind ? String(r.piece_kind) + '：' : '') + (r.error || '失败')
                );
              }
            }
          });

          // 先检测存在 → 弹窗确认是否覆盖 → 再覆盖提交
          if (hasExists && okn === 0) {
            var yes = window.confirm('检测到该签署人已存在同类别元件（按签署人+元件类别判断），是否覆盖？');
            if (!yes) {
              setPieceBatchFeedback('已取消覆盖：队列未清空，你可以调整选择或换签署人后再上传。', true);
              savePieceBatchDraft('upload_cancelled');
              return Promise.resolve();
            }
            return uploadPiecesSequential(true).then(handleUploadResult);
          }

          if (okn === 0 && (failn > 0 || batchQueue.length > 0)) {
            setPieceBatchFeedback(
              '全部未保存成功' +
                (firstErrs.length ? '（' + firstErrs.join('；') + '）' : '') +
                '。请根据提示修正后可从草稿箱重试上传。',
              true
            );
            savePieceBatchDraft('upload_failed');
            setPieceHintFeedback(
              '服务器明细：成功 0 条' + (failn ? '，失败 ' + failn + ' 条' : '') + '。',
              false
            );
            _refreshSignersAfterPieceUpload();
            return Promise.resolve();
          }
          clearFileRegionErr();
          clearPieceBatchDraft();
          setPieceHintFeedback(
            '批量保存完成：成功 ' + okn + ' 条' + (failn ? '，失败 ' + failn + ' 条' : '') + '。',
            false
          );
          resetPieceBatchUi();
          _refreshSignersAfterPieceUpload();
          return Promise.resolve();
        }

        withButtonBusy(pieceBatchUploadBtn, '上传中…', function () {
          savePieceBatchDraft('uploading');
          setPieceBatchFeedback(
            '将逐项上传共 ' +
              batchQueue.length +
              ' 项元件（每项最长约 ' +
              Math.round(PIECE_BATCH_ITEM_TIMEOUT_MS / 1000) +
              's，失败自动重试；可点「取消上传」中止）…',
            false
          );
          return uploadPiecesSequential(false)
            .then(handleUploadResult)
            .finally(function () {
              _endPieceBatchUploadAbort();
            });
        }).catch(function (e) {
          _endPieceBatchUploadAbort();
          if (_isUploadCancelledError(e) || pieceBatchUploadUserCancel || /已取消上传/.test(String(e.message || ''))) {
            pieceBatchUploadUserCancel = false;
            savePieceBatchDraft('upload_cancelled');
            setPieceBatchFeedback('已取消上传。队列仍在，可点「上传队列中全部」重试。', true);
            return;
          }
          savePieceBatchDraft('upload_failed');
          setPieceBatchFeedback(e.message || String(e), true);
          _refreshSignersAfterPieceUpload();
        });
      });

      pieceBatchCancelBtn.addEventListener('click', function () {
        if (_isPieceBatchUploading()) {
          if (!window.confirm('确定取消当前上传？\n\n队列与本地草稿会保留，可稍后点「上传队列中全部」重试。')) {
            return;
          }
          pieceBatchUploadUserCancel = true;
          if (pieceBatchUploadAc) {
            try {
              pieceBatchUploadAc.abort();
            } catch (_) {}
          }
          setPieceBatchFeedback('正在取消上传…', true);
          return;
        }
        if ((batchQueue && batchQueue.length) || batchActive) {
          if (
            !window.confirm(
              '确定取消批量录入？\n\n将清空未上传队列与本地草稿；已保存到服务器的元件不会被删除。'
            )
          ) {
            return;
          }
        }
        resetPieceBatchUi();
        clearFileRegionErr();
        setPieceHintFeedback('', false);
      });
      _setPieceBatchCancelLabel(false);
    })();

    if (IS_FILE_SIGN_PAGE && signedSearchBtn && signedSearchInput) {
      signedSearchBtn.addEventListener('click', function () {
        signedListQ = signedSearchInput.value ? String(signedSearchInput.value).trim() : '';
        signedListPage = 1;
        refreshSignedList();
      });
      signedSearchInput.addEventListener('keydown', function (ev) {
        if (ev.key === 'Enter') {
          ev.preventDefault();
          signedSearchBtn.click();
        }
      });
    }
    if (IS_FILE_SIGN_PAGE && signedPrevBtn) signedPrevBtn.addEventListener('click', function () {
      if (signedListPage <= 1) return;
      signedListPage -= 1;
      refreshSignedList();
    });
    if (IS_FILE_SIGN_PAGE && signedNextBtn) signedNextBtn.addEventListener('click', function () {
      signedListPage += 1;
      refreshSignedList();
    });

    if (IS_MATERIALS_PAGE && strokeItemSearchBtn && strokeItemSearchInput) {
      strokeItemSearchBtn.addEventListener('click', function () {
        strokeItemQ = strokeItemSearchInput.value ? String(strokeItemSearchInput.value).trim() : '';
        strokeItemPage = 1;
        refreshStrokeItemList();
      });
      strokeItemSearchInput.addEventListener('keydown', function (ev) {
        if (ev.key === 'Enter') {
          ev.preventDefault();
          strokeItemSearchBtn.click();
        }
      });
    }
    if (IS_MATERIALS_PAGE && strokeItemCatSelect) strokeItemCatSelect.addEventListener('change', function () {
      strokeItemCat = strokeItemCatSelect.value || '';
      strokeItemPage = 1;
      refreshStrokeItemList();
    });
    if (IS_MATERIALS_PAGE && strokeItemPrevBtn) strokeItemPrevBtn.addEventListener('click', function () {
      if (strokeItemPage <= 1) return;
      strokeItemPage -= 1;
      refreshStrokeItemList();
    });
    if (IS_MATERIALS_PAGE && strokeItemNextBtn) strokeItemNextBtn.addEventListener('click', function () {
      strokeItemPage += 1;
      refreshStrokeItemList();
    });
  }

  function selectedRoleIds() {
    return ROLES.map(function (r) {
      return r.id;
    }).filter(function (id) {
      var el = document.getElementById('chk_' + id);
      return el && el.checked;
    });
  }

  function updateSubmitState() {
    // 素材录入页没有「生成已签名文档」按钮；避免空指针导致脚本中断
    if (!submitBtn) return;
    var wbMode = isBatchWorkbenchMode();
    var batchMode = !!(batchModeCb && batchModeCb.checked);
    var picked = _checkedBatchFileIds().length;
    var src = signSourceValue();
    var libRestrict = libraryRolesRestrictedToChecks();
    var ok = false;
    // 批量工作台：以表格勾选为准，不要求用户再勾「批量模式」复选框
    if (wbMode && picked > 0) {
      ok = !!signersDbShare;
      if (ok && src === 'library' && libRestrict && !selectedRoleIds().length) {
        ok = false;
      }
    } else if (batchMode) {
      ok = !!(signersDbShare && picked > 0);
      if (ok && src === 'library' && libRestrict && !selectedRoleIds().length) {
        ok = false;
      }
    } else if (!selectedFileId) {
      ok = false;
    } else if (src === 'canvas') {
      ok = selectedRoleIds().length > 0;
    } else if (!signersDbShare) {
      ok = false;
    } else if (libRestrict) {
      ok = selectedRoleIds().length > 0;
    } else {
      ok = libraryBoundSignableRoleIds(currentRoleMap).length > 0;
    }
    submitBtn.disabled = !ok;
    if (wbMode && picked > 1) {
      submitBtn.textContent = '批量生成 ' + picked + ' 个已签名文档';
    } else if (batchMode || (wbMode && picked > 0)) {
      submitBtn.textContent = picked > 1 ? '批量生成 ' + picked + ' 个已签名文档' : '生成已签名文档';
    } else {
      submitBtn.textContent = '生成已签名文档';
    }
  }

  /** FTP 优先；失败时 err 来自后端 ftp_last_error（MySQL 保底时仍有下载）。 */
  function ftpMetaLine(uploaded, errMsg, blobNote) {
    if (uploaded === false) {
      var e = (errMsg && String(errMsg).trim()) || '';
      var base = 'FTP：未上传（' + blobNote + '）';
      return e ? base + ' 原因：' + e.slice(0, 200) + (e.length > 200 ? '…' : '') : base;
    }
    if (uploaded === true) return 'FTP：已上传';
    return '';
  }

  function isBatchWorkbenchMode() {
    try {
      return document.body.classList.contains('batch-workbench-mode');
    } catch (_) {
      return false;
    }
  }

  var EXECUTOR_STRONG_LABEL_RE =
    /测试人|测试人员|试验人|试验人员|执行人员|执行人|实施人|实施人员|操作人|操作人员|经手人|Tester|Test\s*Person|Tested\s*by|Test\s*engineer/i;
  var REVIEWER_STRONG_LABEL_RE =
    /复核人|复核人员|审核人|审核人员|Reviewer|Reviewed\s*by|Checked\s*by/i;

  /** detect/映射用 role_id：审核人员/复核人员等与 reviewer 同义 */
  function _canonicalSignRoleId(roleId) {
    var rid = String(roleId || '').trim();
    if (rid === 'reviewer_tail') return 'reviewer';
    return rid;
  }

  function _fileDisplayNameById(fileId) {
    var rec = savedFiles.find(function (x) {
      return x && String(x.id) === String(fileId);
    });
    return rec && rec.name ? String(rec.name) : '';
  }

  function _isUseCaseExecutionTableFileName(fileId) {
    var nm = _fileDisplayNameById(fileId);
    return /用例执行表|用例执行记录/.test(nm);
  }

  /** 文件名是否为约定「用例表」（排除用例执行表） */
  function _isUseCaseSpecTableNoSignByFileName(fileId) {
    var nm = _fileDisplayNameById(fileId);
    if (/用例执行表|用例执行记录/.test(nm)) return false;
    return /用例表/.test(nm);
  }

  /** 测试任务表 / 测试任务执行表：约定无需签字（与用例执行表不同） */
  function _isTestTaskTableNoSignByFileName(fileId) {
    var nm = _fileDisplayNameById(fileId);
    return /测试任务执行表/.test(nm) || /测试任务表/.test(nm);
  }

  /** 约定无需签字的附件表（用例表、测试任务表等） */
  function isNoSignRequiredForFile(fileId, dataOpt) {
    if (_isUseCaseSpecTableNoSignByFileName(fileId)) return true;
    if (_isTestTaskTableNoSignByFileName(fileId)) return true;
    var data = dataOpt;
    if (data === undefined) {
      data = (fileUiCache[String(fileId)] || {}).lastDetectData;
    }
    var rr = data && data.document_role_rule;
    if (
      rr &&
      rr.matched &&
      (rr.category === 'use_case_spec_table' || rr.category === 'test_task_no_sign')
    ) {
      return true;
    }
    return false;
  }

  /** 约定无需签字的表可跳过签字位识别（轻量规则路径） */
  function shouldSkipDetectPipelineForFile(fileId) {
    return _isUseCaseSpecTableNoSignByFileName(fileId) || _isTestTaskTableNoSignByFileName(fileId);
  }

  function _isUseCaseTableFileName(fileId) {
    return _isUseCaseExecutionTableFileName(fileId);
  }

  function _isTestPlanStyleFileName(fileId) {
    var rec = savedFiles.find(function (x) {
      return x && String(x.id) === String(fileId);
    });
    var nm = rec && rec.name ? String(rec.name) : '';
    return /测试方案|测试计划|test\s*plan/i.test(nm);
  }

  function _blockLabelText(b) {
    if (!b) return '';
    return String(b.label_preview || b.source_hint || '');
  }

  function _detectRoleLabelInData(data, re, roleId) {
    if (!data || !data.ok) return false;
    var i;
    var j;
    for (i = 0; i < (data.roles || []).length; i++) {
      if (data.roles[i] && _canonicalSignRoleId(data.roles[i].id) === roleId) return true;
    }
    for (i = 0; i < (data.blocks || []).length; i++) {
      var b = data.blocks[i];
      if (re.test(_blockLabelText(b))) return true;
      var fields = b && b.fields ? b.fields : [];
      for (j = 0; j < fields.length; j++) {
        var f = fields[j];
        if (f && f.type === 'role_id' && _canonicalSignRoleId(f.name) === roleId) return true;
      }
    }
    return false;
  }

  function _detectExecutorViaStrongLabel(data) {
    return _detectRoleLabelInData(data, EXECUTOR_STRONG_LABEL_RE, 'executor');
  }

  function _detectReviewerViaStrongLabel(data) {
    return _detectRoleLabelInData(data, REVIEWER_STRONG_LABEL_RE, 'reviewer');
  }

  function _ensureUseCaseTableRoles(norm, data) {
    var out = norm.slice();
    var seen = {};
    out.forEach(function (r) {
      if (r && r.id) seen[r.id] = true;
    });
    if (!seen.executor && _detectExecutorViaStrongLabel(data)) {
      out.push({ id: 'executor', conf: 0.72 });
      seen.executor = true;
    }
    if (!seen.reviewer && _detectReviewerViaStrongLabel(data)) {
      out.push({ id: 'reviewer', conf: 0.72 });
      seen.reviewer = true;
    }
    return out;
  }

  function _dedupeNormalizedRoles(list) {
    var out = [];
    var seen = {};
    (list || []).forEach(function (r) {
      if (!r || !r.id) return;
      var id = _canonicalSignRoleId(r.id);
      if (seen[id]) return;
      seen[id] = true;
      out.push({ id: id, conf: r.conf });
    });
    return out;
  }

  /** aiword 交接：规范化角色 id；用例执行表页脚「测试人/复核人」须保留执行人员+审核 */
  function _filterAiwordHandoffRoles(out, data, fileId) {
    var norm = _dedupeNormalizedRoles(out);
    if (isNoSignRequiredForFile(fileId, data)) {
      return [];
    }
    if (_isUseCaseExecutionTableFileName(fileId)) {
      return _ensureUseCaseTableRoles(norm, data);
    }
    if (!_isFromAiwordHandoff()) return norm;
    var hasExecutor = false;
    norm.forEach(function (r) {
      if (r && r.id === 'executor') hasExecutor = true;
    });
    var keepExecutor =
      hasExecutor &&
      (_isTestPlanStyleFileName(fileId) || _detectExecutorViaStrongLabel(data));
    if (!hasExecutor) return norm;
    if (keepExecutor) return norm;
    return norm.filter(function (r) {
      return r.id !== 'executor';
    });
  }

  function mergeDetectedRolesFromData(data, fileIdOpt) {
    var out = [];
    var seen = {};
    if (!data || !data.ok) return out;
    (data.roles || []).forEach(function (x) {
      if (!x || !x.id) return;
      var id = _canonicalSignRoleId(x.id);
      if (seen[id]) return;
      seen[id] = true;
      out.push({ id: id, conf: x.confidence });
    });
    (data.blocks || []).forEach(function (b) {
      var bc = b && typeof b.confidence === 'number' ? b.confidence : null;
      var hint = _blockLabelText(b);
      (b && b.fields ? b.fields : []).forEach(function (f) {
        if (!f || f.type !== 'role_id' || !f.name) return;
        var id = _canonicalSignRoleId(f.name);
        if (seen[id]) return;
        seen[id] = true;
        out.push({ id: id, conf: bc });
      });
      if (EXECUTOR_STRONG_LABEL_RE.test(hint) && !seen.executor) {
        seen.executor = true;
        out.push({ id: 'executor', conf: bc });
      }
      if (REVIEWER_STRONG_LABEL_RE.test(hint) && !seen.reviewer) {
        seen.reviewer = true;
        out.push({ id: 'reviewer', conf: bc });
      }
    });
    var fid = fileIdOpt || selectedFileId;
    if (isNoSignRequiredForFile(fid, data)) {
      return [];
    }
    out = _filterAiwordHandoffRoles(out, data, fid);
    if (_isUseCaseExecutionTableFileName(fid)) {
      out = _ensureUseCaseTableRoles(out, data);
    }
    return out;
  }

  function mergeDetectedRolesForFile(fileId) {
    var st = fileUiCache[String(fileId)] || {};
    return mergeDetectedRolesFromData(st.lastDetectData, fileId);
  }

  function _buildDetectEvidenceSummary(data, roleIds) {
    if (!data || !data.ok) return '';
    var lines = [];
    var rr = data.document_role_rule || null;
    if (rr && rr.matched) {
      var pat = rr.pattern || '';
      var rtxt = Array.isArray(rr.roles) && rr.roles.length
        ? rr.roles.map(function (x) { return roleLabel(x); }).join('、')
        : '无需签字';
      lines.push('规则命中：' + pat + ' -> ' + rtxt);
    }
    var dc = data.detect_correction;
    if (dc && typeof dc === 'object' && (dc.wrong_description || (dc.expected_roles && dc.expected_roles.length))) {
      var expTxt = Array.isArray(dc.expected_roles) && dc.expected_roles.length
        ? dc.expected_roles.map(function (x) { return roleLabel(x); }).join('、')
        : '';
      lines.push(
        '人工纠正：' +
          String(dc.wrong_description || '').slice(0, 80) +
          (expTxt ? (' → 应为 ' + expTxt) : '')
      );
    }
    var ds0 = data.debug_summary || null;
    if (ds0 && ds0.correction_override) {
      lines.push('重新识别已应用人工纠正登记');
    }
    var evid = (data.role_evidence && typeof data.role_evidence === 'object') ? data.role_evidence : {};
    (roleIds || []).forEach(function (rid) {
      var arr = Array.isArray(evid[rid]) ? evid[rid] : [];
      if (!arr.length) return;
      var top = arr[0] || {};
      var src = String(top.source_hint || '').trim();
      var rules = Array.isArray(top.matched_rules) ? top.matched_rules.join(',') : '';
      var conf = Number(top.confidence || 0);
      var confText = isFinite(conf) && conf > 0 ? (' @' + conf.toFixed(2)) : '';
      var tail = [];
      if (src) tail.push(src);
      if (rules) tail.push(rules);
      lines.push(roleLabel(rid) + confText + (tail.length ? (' [' + tail.join(' | ') + ']') : ''));
    });
    var ds = data.debug_summary || null;
    if (ds && ds.kind) {
      var bcnt = Number(ds.total_blocks || 0);
      if (isFinite(bcnt) && bcnt >= 0) lines.push('检测块: ' + bcnt + ' (' + ds.kind + ')');
    }
    return lines.join('\n');
  }

  function _detectRoleIdsFromBlock(block) {
    var out = [];
    if (!block || !Array.isArray(block.fields)) return out;
    block.fields.forEach(function (f) {
      if (!f || String(f.type || '') !== 'role_id') return;
      var rid = _canonicalSignRoleId(f.name);
      if (!rid) return;
      if (out.indexOf(rid) < 0) out.push(rid);
    });
    return out;
  }

  function _axisToChineseLabel(axis) {
    switch (axis) {
      case 'horizontal':
        return '从左到右';
      case 'vertical':
        return '从上到下';
      case 'mixed':
        return '混合排列';
      case 'single':
        return '仅 1 个角色';
      case 'inline':
        return '段落内联';
      default:
        return '';
    }
  }

  function _relationToChineseLabel(rel) {
    switch (rel) {
      case 'same_cell':
        return '同一单元格';
      case 'different_cell':
        return '不同单元格';
      case 'paragraph_inline':
        return '正文段落';
      case 'none':
        return '';
      default:
        return '';
    }
  }

  function _positionToChineseLabel(pos) {
    switch (pos) {
      case 'right':
        return '日期在角色右方';
      case 'below':
        return '日期在角色下方';
      case 'inline':
        return '与角色同行';
      case 'none':
        return '';
      default:
        return '';
    }
  }

  function _separatorToChineseLabel(sep) {
    if (!sep) return '';
    if (sep === 'slash') return '/';
    if (sep === 'backslash') return '\\';
    if (sep === 'space') return '空格';
    if (sep === 'newline') return '换行';
    if (sep === 'empty_cell') return '空单元格';
    if (sep === 'cell') return '单元格';
    if (sep === 'adjacent') return '紧邻无分隔';
    if (sep === 'none') return '';
    if (sep === 'unknown') return '';
    if (sep.indexOf && sep.indexOf('punct:') === 0) return sep.slice(6);
    if (sep.indexOf && sep.indexOf('other:') === 0) return sep.slice(6);
    return sep;
  }

  function _pickBestBlockForRole(blocks, rid) {
    var best = null;
    var bestKey = null;
    (blocks || []).forEach(function (b) {
      if (!b || typeof b !== 'object') return;
      var bRoles = _detectRoleIdsFromBlock(b).filter(function (r) {
        return r === rid;
      });
      if (!bRoles.length) return;
      var preview = String(b.label_preview || '');
      var key = (preview ? preview.length : 9999) + (1 - Number(b.confidence || 0));
      if (best === null || key < bestKey) {
        best = b;
        bestKey = key;
      }
    });
    return best;
  }

  function _layoutFromDetectBlock(block) {
    if (!block) {
      return {
        name_slot: false,
        date_slot: false,
        date_relation: 'none',
        date_position: 'none',
        separator: 'none',
      };
    }
    var preview = String(block.label_preview || '');
    var matched = Array.isArray(block.matched_rules)
      ? block.matched_rules.map(function (x) {
          return String(x || '');
        })
      : [];
    var fields = Array.isArray(block.fields) ? block.fields : [];
    var hasDate = fields.some(function (f) {
      return f && String(f.type || '') === 'date';
    });
    if (/\/|／/.test(preview)) {
      return {
        name_slot: true,
        date_slot: true,
        date_relation: 'same_cell',
        date_position: 'right',
        separator: 'slash',
      };
    }
    if (/\r|\n/.test(preview)) {
      return {
        name_slot: true,
        date_slot: hasDate,
        date_relation: 'same_cell',
        date_position: 'below',
        separator: 'newline',
      };
    }
    if (matched.indexOf('docx_role_with_date_row') >= 0) {
      return {
        name_slot: true,
        date_slot: hasDate,
        date_relation: 'different_cell',
        date_position: 'below',
        separator: 'cell',
      };
    }
    if (preview.indexOf('|') >= 0 && hasDate) {
      return {
        name_slot: true,
        date_slot: true,
        date_relation: 'different_cell',
        date_position: 'right',
        separator: 'cell',
      };
    }
    if (hasDate && /\s+\d{4}/.test(preview)) {
      return {
        name_slot: true,
        date_slot: true,
        date_relation: 'same_cell',
        date_position: 'right',
        separator: 'space',
      };
    }
    if (hasDate) {
      return {
        name_slot: true,
        date_slot: true,
        date_relation: 'different_cell',
        date_position: 'right',
        separator: 'cell',
      };
    }
    return {
      name_slot: true,
      date_slot: false,
      date_relation: 'none',
      date_position: 'none',
      separator: 'none',
    };
  }

  function _inferArrangementFromBlocks(blocks, ids) {
    var rowByRole = {};
    var multiRoleBlock = null;
    (blocks || []).forEach(function (b) {
      if (!b) return;
      var bRoles = _detectRoleIdsFromBlock(b).filter(function (rid) {
        return ids.indexOf(rid) >= 0;
      });
      if (bRoles.length >= 2) multiRoleBlock = b;
      bRoles.forEach(function (rid) {
        var m = /table\d+\.row(\d+)/i.exec(String(b.source_hint || ''));
        if (m) rowByRole[rid] = parseInt(m[1], 10);
      });
    });
    var rows = ids
      .map(function (rid) {
        return rowByRole[rid];
      })
      .filter(function (x) {
        return x != null;
      });
    if (rows.length >= 2) {
      var uniq = {};
      rows.forEach(function (r) {
        uniq[r] = true;
      });
      if (Object.keys(uniq).length === 1) return 'horizontal';
      return 'vertical';
    }
    if (multiRoleBlock) {
      var prev = String(multiRoleBlock.label_preview || '');
      if (prev.indexOf('|') >= 0) return 'horizontal';
      if (/\r|\n/.test(prev)) return 'vertical';
    }
    return '';
  }

  function _inferSlotLayoutFromBlocks(data, ids) {
    var blocks = Array.isArray(data && data.blocks) ? data.blocks : [];
    var role_layouts = {};
    ids.forEach(function (rid) {
      var b = _pickBestBlockForRole(blocks, rid);
      if (b) role_layouts[rid] = _layoutFromDetectBlock(b);
    });
    var arrangement = _inferArrangementFromBlocks(blocks, ids) || 'unknown';
    return { arrangement: arrangement, role_layouts: role_layouts };
  }

  function _metricsFromExpectedSlotLayout(esl, ids) {
    if (!esl || typeof esl !== 'object') return null;
    var rel = String(esl.date_relation || '').trim();
    var pos = String(esl.date_position || '').trim();
    var sep = String(esl.separator || '').trim();
    var axis = String(esl.arrangement || '').trim() || 'unknown';
    var sameCell = esl.same_cell;
    if (sameCell === true || rel === 'same_cell') sameCell = true;
    else if (sameCell === false || rel === 'different_cell') sameCell = false;
    else sameCell = null;
    var role_layouts = {};
    ids.forEach(function (rid) {
      role_layouts[rid] = {
        name_slot: true,
        date_slot: rel && rel !== 'none',
        date_relation: rel || 'none',
        date_position: pos || 'none',
        separator: sep || 'none',
      };
    });
    return {
      arrangement: axis,
      arrangementLabel: _axisToChineseLabel(axis) || '未判定',
      dateRelationLabel: _relationToChineseLabel(rel) || '未判定',
      datePositionLabel: _positionToChineseLabel(pos) || '未判定',
      separatorLabel: _separatorToChineseLabel(sep) || '未判定',
      sameCellLabel:
        sameCell === true ? '是' : sameCell === false ? '否' : '未判定',
      role_layouts: role_layouts,
      source: 'correction',
    };
  }

  /** 汇总签字位版式：优先 signature_layout → 识别块推断 → 人工登记 */
  function _computeSlotLayoutMetrics(data, ids) {
    var layout =
      data && data.signature_layout && typeof data.signature_layout === 'object'
        ? data.signature_layout
        : null;
    var role_layouts = {};
    if (layout && layout.role_layouts && typeof layout.role_layouts === 'object') {
      role_layouts = layout.role_layouts;
    }
    var axis = layout && layout.arrangement ? String(layout.arrangement) : '';
    ids.forEach(function (rid) {
      if (!role_layouts[rid]) {
        var fb = _inferSlotLayoutFromBlocks(data, [rid]);
        if (fb.role_layouts[rid]) role_layouts[rid] = fb.role_layouts[rid];
      }
    });
    if (!axis || axis === 'unknown') {
      var fbAll = _inferSlotLayoutFromBlocks(data, ids);
      axis = fbAll.arrangement || axis;
      ids.forEach(function (rid) {
        if (!role_layouts[rid] && fbAll.role_layouts[rid]) {
          role_layouts[rid] = fbAll.role_layouts[rid];
        }
      });
    }
    var corrEsl =
      data &&
      data.detect_correction &&
      data.detect_correction.expected_slot_layout;
    var corrMetrics = _metricsFromExpectedSlotLayout(corrEsl, ids);
    if (corrMetrics) {
      role_layouts = corrMetrics.role_layouts;
      axis = corrMetrics.arrangement;
    }
    var rels = {};
    var posSet = {};
    var seps = {};
    ids.forEach(function (rid) {
      var info = role_layouts[rid] || {};
      if (info.date_relation && info.date_relation !== 'none') {
        rels[info.date_relation] = true;
      }
      if (info.date_position && info.date_position !== 'none') {
        posSet[info.date_position] = true;
      }
      if (info.separator && info.separator !== 'none') {
        seps[info.separator] = true;
      }
    });
    var relKeys = Object.keys(rels);
    var posKeys = Object.keys(posSet);
    var sepKeys = Object.keys(seps);
    var sameCellLabel = '未判定';
    if (relKeys.length === 1 && relKeys[0] === 'same_cell') sameCellLabel = '是';
    else if (relKeys.length && relKeys.indexOf('same_cell') < 0) sameCellLabel = '否';
    else if (relKeys.length > 1) sameCellLabel = '混合';
    if (corrMetrics && corrMetrics.sameCellLabel !== '未判定') {
      sameCellLabel = corrMetrics.sameCellLabel;
    }
    return {
      arrangement: axis || 'unknown',
      arrangementLabel: _axisToChineseLabel(axis) || '未判定',
      dateRelationLabel:
        relKeys.length === 1
          ? _relationToChineseLabel(relKeys[0])
          : relKeys.length > 1
            ? '混合'
            : '未判定',
      datePositionLabel:
        posKeys.length === 1
          ? _positionToChineseLabel(posKeys[0])
          : posKeys.length > 1
            ? '混合'
            : '未判定',
      separatorLabel:
        sepKeys.length === 1
          ? _separatorToChineseLabel(sepKeys[0])
          : sepKeys.length > 1
            ? '混合'
            : '未判定',
      sameCellLabel: sameCellLabel,
      role_layouts: role_layouts,
      source: corrMetrics ? 'correction' : layout && layout.ok ? 'signature_layout' : 'blocks',
    };
  }

  function _buildLayoutSentence(ids, axis, role_layouts) {
    if (!role_layouts || !Object.keys(role_layouts).length) return '';
    var roleNames = ids
      .map(function (rid) { return roleLabel(rid); })
      .filter(Boolean);
    var axisLabel = _axisToChineseLabel(axis);
    var relSet = {};
    var posSet = {};
    var sepSet = {};
    ids.forEach(function (rid) {
      var info = role_layouts[rid];
      if (!info) return;
      if (info.date_relation && info.date_relation !== 'none') relSet[info.date_relation] = true;
      if (info.date_position && info.date_position !== 'none') posSet[info.date_position] = true;
      if (info.separator && info.separator !== 'none') sepSet[info.separator] = true;
    });
    var relKeys = Object.keys(relSet);
    var posKeys = Object.keys(posSet);
    var sepKeys = Object.keys(sepSet);
    var rel = relKeys.length === 1 ? _relationToChineseLabel(relKeys[0]) : (relKeys.length > 1 ? '存在多种排布' : '');
    var pos = posKeys.length === 1 ? _positionToChineseLabel(posKeys[0]) : (posKeys.length > 1 ? '位置不统一' : '');
    var sep = sepKeys.length === 1 ? _separatorToChineseLabel(sepKeys[0]) : (sepKeys.length > 1 ? '混合' : '');
    var parts = [];
    if (roleNames.length) parts.push('需签角色为' + roleNames.join('、'));
    if (axisLabel) parts.push(axisLabel + '排列');
    if (rel) parts.push('角色与日期' + rel);
    if (pos) parts.push(pos);
    if (sep) parts.push('分隔方式为' + sep);
    return parts.join('，');
  }

  function _buildSignSlotSummary(data, roleIds) {
    if (!data || !data.ok) {
      return { label: '—', detail: '', tags: [] };
    }
    var ids = [];
    (roleIds || []).forEach(function (rid) {
      var id = _canonicalSignRoleId(rid);
      if (!id || ids.indexOf(id) >= 0) return;
      ids.push(id);
    });
    if (!ids.length) {
      return { label: '—', detail: '', tags: [] };
    }

    var layoutMetrics = _computeSlotLayoutMetrics(data, ids);
    var role_layouts = layoutMetrics.role_layouts || {};
    var axis = layoutMetrics.arrangement || 'unknown';

    var blockHasDateByRole = {};
    var sourceHints = {};
    var previews = [];
    var blocks = Array.isArray(data.blocks) ? data.blocks : [];
    blocks.forEach(function (b) {
      if (!b || typeof b !== 'object') return;
      var bRoles = _detectRoleIdsFromBlock(b).filter(function (rid) {
        return ids.indexOf(rid) >= 0;
      });
      if (!bRoles.length) return;
      var fields = Array.isArray(b.fields) ? b.fields : [];
      var hasDate = fields.some(function (f) {
        return f && String(f.type || '') === 'date';
      });
      var hint = String(b.source_hint || '').trim();
      var preview = String(b.label_preview || '').trim();
      if (preview && previews.indexOf(preview) < 0) previews.push(preview);
      bRoles.forEach(function (rid) {
        if (hasDate) blockHasDateByRole[rid] = true;
        if (hint) {
          if (!sourceHints[rid]) sourceHints[rid] = [];
          if (sourceHints[rid].indexOf(hint) < 0) sourceHints[rid].push(hint);
        }
      });
    });

    var probe = data.slot_probe && typeof data.slot_probe === 'object' ? data.slot_probe : null;
    var probeMissingSet = {};
    if (probe && Array.isArray(probe.missing_roles)) {
      probe.missing_roles.forEach(function (rid) {
        probeMissingSet[String(rid)] = true;
      });
    }
    var perRoleProbe = probe && probe.per_role_results && typeof probe.per_role_results === 'object'
      ? probe.per_role_results
      : {};

    var stat = {};
    var nameOk = 0;
    var dateOk = 0;
    var missingNameRoles = [];
    var missingDateRoles = [];
    ids.forEach(function (rid) {
      var info = role_layouts[rid] || {};
      var hasName = !!info.name_slot;
      var hasDate = !!info.date_slot;
      if (!hasName && blockHasDateByRole[rid]) {
        // 角色在 block 中出现但 layout 没解析出 cell，仍记为「有姓名位但版式未知」
        hasName = true;
      }
      if (!hasDate && blockHasDateByRole[rid]) {
        // layout 未解析出日期位时，回退到 detect blocks 的 date 证据，避免“全部日期位丢失”
        hasDate = true;
      }
      var probeOne = perRoleProbe[rid] && typeof perRoleProbe[rid] === 'object'
        ? perRoleProbe[rid]
        : null;
      var probePlaced = !!(probeOne && probeOne.placed && !probeMissingSet[rid]);
      if (!hasName && probePlaced) {
        // 真实落位探测已成功，姓名位至少可落，避免仅因版式解析缺字段而误报「缺姓名位」。
        hasName = true;
        if (!info.name_loc) {
          info = Object.assign({}, info, { name_loc: 'slot_probe' });
        }
      }
      if (!hasDate && probePlaced) {
        // 日期位仅在已有“日期关系证据”或 block 日期证据时回填，避免过度乐观。
        var rel = String(info.date_relation || '').trim();
        if (blockHasDateByRole[rid] || rel === 'same_cell' || rel === 'different_cell') {
          hasDate = true;
          if (!info.date_loc) {
            info = Object.assign({}, info, { date_loc: 'slot_probe' });
          }
        }
      }
      stat[rid] = {
        hasNameSlot: hasName,
        hasDateSlot: hasDate,
        info: info,
      };
      if (hasName) nameOk++;
      else missingNameRoles.push(roleLabel(rid));
      if (hasDate) dateOk++;
      else missingDateRoles.push(roleLabel(rid));
    });

    var probeOk = probe ? !!probe.ok : null;

    var tags = [];
    var statusText = '待确认';
    var badgeClass = 'warn';
    if (missingNameRoles.length || missingDateRoles.length) {
      // 真实结构分析显示有缺位 → 一定不能标记可签
      if (missingNameRoles.length && missingDateRoles.length) {
        statusText = '缺姓名+日期位';
        badgeClass = 'fail';
      } else if (missingNameRoles.length) {
        statusText = '缺姓名位';
        badgeClass = 'fail';
      } else {
        statusText = '缺日期位';
        badgeClass = 'warn';
      }
    } else if (probeOk === false) {
      statusText = '不可签';
      badgeClass = 'fail';
    } else if (probeOk === true) {
      statusText = '可签';
      badgeClass = 'ok';
    } else {
      statusText = '位齐全';
      badgeClass = 'ok';
    }

    if (missingNameRoles.length) tags.push('姓名位缺失');
    if (missingDateRoles.length) tags.push('日期位缺失');
    if (statusText && statusText !== '—') tags.push(statusText);
    if (statusText === '可签') tags.push('可签');
    if (probeOk === false) tags.push('不可签');
    if (axis === 'horizontal') tags.push('版式-左右');
    else if (axis === 'vertical') tags.push('版式-上下');
    else if (axis === 'mixed') tags.push('版式-混合');
    ids.forEach(function (rid) {
      var info = role_layouts[rid] || {};
      if (info.date_relation === 'same_cell') tags.push('版式-同格');
      if (info.date_relation === 'different_cell') tags.push('版式-分格');
      if (info.date_relation === 'paragraph_inline') tags.push('版式-正文');
      if (info.separator === 'slash') tags.push('分隔-/');
      if (info.separator === 'space') tags.push('分隔-空格');
      if (info.separator === 'empty_cell') tags.push('分隔-空格子');
      if (info.separator === 'cell') tags.push('分隔-单元格');
      if (info.separator === 'newline') tags.push('分隔-换行');
    });

    // 自然语句版式描述（按用户口径）
    var sentence = _buildLayoutSentence(ids, axis, role_layouts);
    var layoutLine = sentence;

    // 列表中只显示精简标签 chip：每角色「名✓/日✓」
    var roleMarks = ids
      .map(function (rid) {
        var one = stat[rid] || {};
        var nm = roleLabel(rid).replace(/人员/g, '');
        if (nm.length > 3) nm = nm.slice(0, 3);
        var mark = one.hasNameSlot && one.hasDateSlot
          ? '名✓ 日✓'
          : one.hasNameSlot
            ? '名✓ 日×'
            : one.hasDateSlot
              ? '名× 日✓'
              : '名× 日×';
        return nm + '(' + mark + ')';
      })
      .join('  ');

    // 详细 tooltip（按“结论/版式/角色明细/校验/样例”分组）
    var detail = [];
    detail.push('【结论】' + statusText + '（姓名位 ' + nameOk + '/' + ids.length + '，日期位 ' + dateOk + '/' + ids.length + '）');
    if (sentence) detail.push('【版式】' + sentence);
    detail.push(
      '【排列方式】' +
        (layoutMetrics.arrangementLabel || '未判定') +
        '（编制/审核/批准/执行/审核人员相对位置）'
    );
    detail.push(
      '【角色与日期】' + (layoutMetrics.dateRelationLabel || '未判定')
    );
    detail.push('【日期位置】' + (layoutMetrics.datePositionLabel || '未判定'));
    detail.push('【分隔方式】' + (layoutMetrics.separatorLabel || '未判定'));
    detail.push('【同一单元格】' + (layoutMetrics.sameCellLabel || '未判定'));
    if (layoutMetrics.source === 'blocks') {
      detail.push('【版式来源】识别块推断（未拿到表格结构分析时）');
    } else if (layoutMetrics.source === 'correction') {
      detail.push('【版式来源】人工登记纠正');
    }
    ids.forEach(function (rid) {
      var one = stat[rid] || {};
      var info = one.info || {};
      var l = roleLabel(rid) + '：姓名位' + (one.hasNameSlot ? '✓' : '×')
        + '，日期位' + (one.hasDateSlot ? '✓' : '×');
      if (info.name_loc) l += '；姓名@' + info.name_loc;
      if (info.date_loc) l += '；日期@' + info.date_loc;
      var sepLabelOne = _separatorToChineseLabel(info.separator);
      if (sepLabelOne) l += '；分隔=' + sepLabelOne;
      if (!info.date_slot && one.hasDateSlot) l += '；日期位来源=识别块证据';
      if (sourceHints[rid] && sourceHints[rid].length) {
        l += '；来源=' + sourceHints[rid].slice(0, 2).join(' / ');
      }
      detail.push(l);
    });
    if (probe && probe.ok === false) {
      var miss = Array.isArray(probe.missing_roles) ? probe.missing_roles : [];
      detail.push(
        '【落位校验】未通过' +
          (miss.length ? '：' + miss.map(function (x) { return roleLabel(x); }).join('、') : '') +
          (probe.error ? '\n' + String(probe.error) : '')
      );
    } else if (probe && probe.ok) {
      detail.push('【落位校验】通过');
    }
    var layoutRaw =
      data.signature_layout && typeof data.signature_layout === 'object'
        ? data.signature_layout
        : null;
    if (layoutRaw && layoutRaw.ok === false && layoutRaw.error) {
      detail.push('【结构分析】失败：' + layoutRaw.error);
    }
    if (previews.length) {
      detail.push('【样例】' + previews.slice(0, 2).join(' ｜ '));
    }

    return {
      label: statusText,
      rolesLine: roleMarks,
      layoutLine: layoutLine,
      badgeClass: badgeClass,
      detail: detail.join('\n'),
      tags: tags,
    };
  }

  /** 仅「流水线进行中」才在签字位列显示处理中（勿用宽泛的 /识别/，否则会误伤「未识别到签字位」「识别失败」等终态） */
  function _isWorkbenchRowPipelineBusy(st) {
    var t = String(st || '').trim();
    if (!t) return false;
    if (t === '识别中…' || t === '识别重试中…' || t === '匹配素材…') return true;
    if (/^识别较慢/.test(t)) return true;
    if (/^匹配素材/.test(t)) return true;
    if (/^分析中/.test(t)) return true;
    return false;
  }

  function _renderWorkbenchSlotCell(td, row) {
    if (!td || !row) return;
    td.className = 'col-slot';
    var st = String(row.status || '');
    if (row._wbProcessing || _isWorkbenchRowPipelineBusy(st)) {
      td.innerHTML = '<span class="wb-slot-badge muted">处理中</span>';
      td.title = '请查看「状态」列';
      return;
    }
    if (row.slotLabel === '无需签字' || (row.slotTags && row.slotTags.indexOf('无需签字') >= 0)) {
      td.innerHTML = '<span class="wb-slot-badge muted">无需签字</span>';
      td.title = row.slotExplain || '规则判定本文件无需签字';
      return;
    }
    var badge = row.slotBadgeClass || 'muted';
    var txt = row.slotLabel || '—';
    var html = '<span class="wb-slot-badge ' + badge + '">' + txt + '</span>';
    if (row.slotRolesLine) {
      html += '<div class="wb-slot-roles">' + row.slotRolesLine + '</div>';
    }
    if (row.slotLayoutLine) {
      html += '<div class="wb-slot-layout">' + row.slotLayoutLine + '</div>';
    }
    td.innerHTML = html;
    td.title = row.slotExplain || txt;
  }

  function validateDetectResponseForFile(fileId, j) {
    if (!j || !j.ok) return '';
    if (j.file_id && String(j.file_id) !== String(fileId)) {
      return '识别结果与当前文件ID不一致';
    }
    var rec = savedFiles.find(function (x) {
      return x && String(x.id) === String(fileId);
    });
    var expectName = rec && rec.name ? _fileBaseNameOnly(rec.name) : '';
    var gotName = j.source_name ? _fileBaseNameOnly(j.source_name) : '';
    if (expectName && gotName && expectName !== gotName) {
      // 实际环境中同一源文件可能经过 .doc/.xls 转换、重命名或数据库展示名修正，
      // 文件名不一致不应直接判定 detect 失败，否则批量会出现“全部识别失败”假象。
      if (!isBatchWorkbenchMode()) {
        return '';
      }
    }
    return '';
  }

  function syncWorkbenchRowFromDetect(fileId, dataOpt) {
    var row = __batchWorkbenchRows[String(fileId)];
    if (!row) return;
    var data = dataOpt;
    if (data === undefined) {
      data = (fileUiCache[String(fileId)] || {}).lastDetectData;
    }
    // 关键：必须传 fileId，否则 _filterAiwordHandoffRoles 用错文件名判断「用例表」
    // 会导致 executor 在多文件批量场景被错误过滤掉。
    var roles = mergeDetectedRolesFromData(data, fileId);
    row.detectedRoleIds = roles.map(function (r) {
      return r.id;
    });
    row.rolesLabel = roles.length
      ? roles
          .map(function (r) {
            return r && r.id ? roleLabel(r.id) : '';
          })
          .filter(Boolean)
          .join('、')
      : '—';
    row.detectExplain = _buildDetectEvidenceSummary(
      data,
      roles.map(function (x) { return x && x.id; }).filter(Boolean)
    );
    var slotData = data;
    if (slotData) {
      var corrCached = __fileDetectCorrectionCache[String(fileId)];
      if (corrCached && !slotData.detect_correction) {
        slotData = Object.assign({}, slotData, { detect_correction: corrCached });
      }
    }
    var slotSummary = _buildSignSlotSummary(
      slotData,
      roles.map(function (x) { return x && x.id; }).filter(Boolean)
    );
    row.slotLabel = slotSummary.label || '—';
    row.slotRolesLine = slotSummary.rolesLine || '';
    row.slotLayoutLine = slotSummary.layoutLine || '';
    row.slotBadgeClass = slotSummary.badgeClass || 'muted';
    row.slotExplain = slotSummary.detail || '';
    row.slotTags = Array.isArray(slotSummary.tags) ? slotSummary.tags.slice() : [];
    var probe = data && data.slot_probe && typeof data.slot_probe === 'object' ? data.slot_probe : null;
    row.slotProbeOk = probe ? !!probe.ok : null;
    // 结构层面：姓名位/日期位是否全部识别到
    row.slotMissingName = row.slotTags.indexOf('姓名位缺失') >= 0;
    row.slotMissingDate = row.slotTags.indexOf('日期位缺失') >= 0;
    // 综合判定：只有真的能在文档落位、且姓名/日期位齐全才算可签
    row.slotSignable =
      row.slotProbeOk === true && !row.slotMissingName && !row.slotMissingDate;
    if (data && data.content_sha256) {
      row.contentSha256 = String(data.content_sha256);
    }
    if (isNoSignRequiredForFile(fileId, data)) {
      row.detectedRoleIds = [];
      row.rolesLabel = '无需签字';
      row.slotLabel = '无需签字';
      row.slotRolesLine = '';
      row.slotLayoutLine = '';
      row.slotBadgeClass = 'muted';
      row.slotExplain = '规则标记该文档无需签字。';
      row.slotTags = ['无需签字'];
      row.slotProbeOk = true;
    }
    var rs = data && data.rule_sync;
    var syncHint = _formatRuleSyncHint(rs);
    if (syncHint) {
      row.detectExplain = row.detectExplain
        ? row.detectExplain + '\n【规则同步】' + syncHint
        : '【规则同步】' + syncHint;
    }
  }

  function _formatRuleSyncHint(rs) {
    if (!rs || !rs.ok) return '';
    var parts = [];
    if (rs.pattern) {
      var verb = rs.action === 'created' ? '已新增' : '已更新';
      parts.push(
        verb +
          '角色规则 `' +
          rs.pattern +
          '` → sign_document_role_rules.json / signature_role_results_T2.md'
      );
    }
    var slotRs = rs.slot_rule_sync;
    if (slotRs && slotRs.ok && slotRs.pattern) {
      var slotVerb = slotRs.action === 'created' ? '已新增' : '已更新';
      parts.push(
        slotVerb +
          '签字位规则 `' +
          slotRs.pattern +
          '` → sign_slot_layout_rules.json / signature_slot_layout_document_rules_T2.md'
      );
    } else if (rs.slot_md_exported) {
      parts.push('已导出 signature_slot_layout_document_rules_T2.md');
    }
    return parts.length ? parts.join('；') : '';
  }

  function roleLibrarySigReady(pair) {
    return !!(pair && pair.sig);
  }

  function roleLibraryDateReady(pair) {
    if (!pair || typeof pair !== 'object') return false;
    if (isCompositeDateMode(pair.date_mode)) {
      return !!(pair.date_iso && !pair.sig);
    }
    return !!(pair.date && !pair.sig);
  }

  /** 库素材「齐全」：有签名，且拼接模式另有 date_iso / 整张日期图 */
  function roleLibraryMaterialReady(pair) {
    if (!pair || typeof pair !== 'object') return false;
    if (!roleLibrarySigReady(pair)) return false;
    if (isCompositeDateMode(pair.date_mode)) {
      return !!(pair.date_iso || pair.date);
    }
    return true;
  }

  function isInternalCacheFileName(name) {
    var base = _fileBaseNameOnly(name);
    if (!base) return false;
    if (/^_(?:ftptpl|dbtpl|link)_/i.test(base)) return true;
    if (/^_ftptpl_[0-9a-f\-]{36}(?:_\d{14})?\.(docx|doc|xlsx|xls)$/i.test(base)) return true;
    return false;
  }

  function repairDisplayFileName(name) {
    var s = String(name || '').trim();
    if (!s) return s;
    if (isInternalCacheFileName(s)) return s;
    if (/[\u4e00-\u9fff]/.test(s)) return s;
    if (!/[\u0080-\u00ff]/.test(s)) return s;
    try {
      var bytes = new Uint8Array(s.length);
      for (var i = 0; i < s.length; i++) bytes[i] = s.charCodeAt(i) & 0xff;
      var dec = new TextDecoder('utf-8').decode(bytes);
      if (dec && /[\u4e00-\u9fff]/.test(dec)) return dec;
    } catch (_) {}
    return s;
  }

  function getFileRoleMapForWorkbench(fileId) {
    // 批量工作台：各行必须读各自 fileUiCache，避免处理 file B 时 selectedFileId=B
    // 导致表格里 file A 的行误用 currentRoleMap（正在写入的 B）而显示不一致。
    if (isBatchWorkbenchMode()) {
      return (fileUiCache[String(fileId)] || {}).currentRoleMap || {};
    }
    if (String(selectedFileId) === String(fileId)) {
      return currentRoleMap || {};
    }
    return (fileUiCache[String(fileId)] || {}).currentRoleMap || {};
  }

  /** 批量工作台：合并 row / 单文件 handoff / 筛选框 中的编审批姓名，保证同一人跨文件用同一套规则匹配。
   *  特别地：当 roleId === 'executor' 且没有专属 hint 时，按业务需求回退到「编写人员」(author) 的 hint，
   *  这样用例表等含「执行人员/测试人员」的文档也能由编写人的签名素材完成签字。 */
  function _workbenchNameHintForRole(fileId, row, roleId) {
    var rid = String(roleId || '').trim();
    if (!rid) return '';
    var h = _workbenchRoleNameHint(row, rid);
    if (h) return h;
    var ctx = __aiwordHandoffCtxByFileId[String(fileId)] || {};
    if (rid === 'author') {
      h = String(ctx.editor || ctx.writer || '').trim();
    } else if (rid === 'reviewer') {
      h = String(ctx.reviewer || '').trim();
    } else if (rid === 'approver') {
      h = String(ctx.approver || '').trim();
    } else if (rid === 'executor') {
      // 执行人员：优先 ctx.executor / row.executor；缺时直接复用编写人员姓名。
      h = String(ctx.executor || ctx.editor || ctx.writer || '').trim();
    }
    if (h) return h;
    try {
      var fq = _ensureRoleFilterStateForFile(fileId, rid);
      if (fq && fq.sig) return String(fq.sig).trim();
    } catch (_) {}
    // 兜底：executor 仍无 hint 时，再尝试 author 已有的筛选/匹配姓名
    if (rid === 'executor') {
      try {
        var authorFq = _ensureRoleFilterStateForFile(fileId, 'author');
        if (authorFq && authorFq.sig) return String(authorFq.sig).trim();
      } catch (_) {}
      // 最后兜底：复用 row 上的 author 姓名（即 row.editor）
      if (row && row.editor) return String(row.editor).trim();
    }
    return '';
  }

  /**
   * 按签署人绑定库素材：签名与日期分别尝试，互不要求同时存在。
   * 有则写入 pair，无则保留 pair 里已有字段（不整行清空）。
   */
  function _applyWorkbenchSignerMaterialToPair(p, signer, strokeLoc, docDateIso) {
    var out = { applied: false, hasSig: false, hasDate: false };
    if (!p || !signer) return out;
    var pl = strokeLoc === 'en' || strokeLoc === 'zh' ? strokeLoc : 'zh';

    var sigId = _firstStrokeId(signer, 'sig', pl);
    if (!sigId && pl !== 'zh') {
      sigId = _firstStrokeId(signer, 'sig', 'zh');
    }
    if (sigId) {
      p.sig = sigId;
      out.hasSig = true;
      out.applied = true;
      var sigLoc = _strokeItemLocale('sig', sigId) || pl;
      if (signersDbShare) {
        p.date_mode = sigLoc === 'en' ? 'composite_en_space' : 'composite_zh_ymd';
      }
    }

    if (signersDbShare) {
      if (docDateIso) {
        p.date_iso = docDateIso;
        p.date = null;
        out.hasDate = true;
        out.applied = true;
        if (!p.date_mode) {
          p.date_mode = pl === 'en' ? 'composite_en_space' : 'composite_zh_ymd';
        }
      }
    } else {
      var dItemId = _firstStrokeId(signer, 'date', pl);
      if (!dItemId && pl !== 'zh') {
        dItemId = _firstStrokeId(signer, 'date', 'zh');
      }
      if (dItemId) {
        p.date = dItemId;
        p.date_mode = 'item';
        out.hasDate = true;
        out.applied = true;
      }
    }
    return out;
  }

  /** 汇总本文件各角色素材缺口（用于工作台提示，不阻断部分匹配结果展示） */
  function _workbenchMatchGapLines(fileId, row) {
    var lines = [];
    var map = getFileRoleMapForWorkbench(fileId);
    var gapRoles = ['author', 'reviewer', 'approver'];
    // 与 _workbenchPlanRoleMap 一致：detect 命中过 executor 时，也展示其缺口提示
    try {
      mergeDetectedRolesForFile(fileId).forEach(function (r) {
        if (r && r.id === 'executor' && gapRoles.indexOf('executor') < 0) {
          gapRoles.push('executor');
        }
      });
    } catch (_) {}
    gapRoles.forEach(function (rid) {
      var hint = _workbenchNameHintForRole(fileId, row, rid);
      var p = map[rid];
      if (!hint && !roleMapEntryNonEmpty(p)) return;
      var label = roleLabel(rid);
      if (hint && !_findSignerByNameHint(hint)) {
        lines.push(label + '：库中未找到签署人「' + hint + '」');
        return;
      }
      if (!roleMapEntryNonEmpty(p)) {
        if (hint) lines.push(label + '（' + hint + '）：未匹配到任何素材');
        return;
      }
      if (!roleLibrarySigReady(p)) {
        lines.push(label + (hint ? '（' + hint + '）' : '') + '：缺签名素材');
      }
      if (!roleLibraryDateReady(p) && !roleLibraryMaterialReady(p)) {
        if (isCompositeDateMode(p.date_mode) && !p.date_iso) {
          lines.push(label + (hint ? '（' + hint + '）' : '') + '：已有签名，缺文档体现日期');
        } else if (!isCompositeDateMode(p.date_mode) && !p.date) {
          lines.push(label + (hint ? '（' + hint + '）' : '') + '：缺日期素材');
        }
      }
    });
    return lines;
  }

  /** 稳定比较 role-map，避免 null/字段顺序导致反复判定 changed 并循环 PUT */
  function _roleMapStableJson(m) {
    var out = {};
    Object.keys(m || {})
      .sort()
      .forEach(function (rid) {
        var p = m[rid];
        if (!p || typeof p !== 'object' || !roleMapEntryNonEmpty(p)) return;
        var q = {};
        if (p.sig) q.sig = String(p.sig);
        if (p.date) q.date = String(p.date);
        if (p.date_mode) q.date_mode = String(p.date_mode);
        if (p.date_iso) q.date_iso = String(p.date_iso);
        out[rid] = q;
      });
    return JSON.stringify(out);
  }

  var _lastPutTsByFile = {};

  function persistWorkbenchRoleMapChange(fileId, nextMap, opt) {
    opt = opt && typeof opt === 'object' ? opt : {};
    var key = String(fileId || '');
    var stable = _roleMapStableJson(nextMap);
    if (!opt.force) {
      if (_lastPersistedRoleMapStable[key] === stable) {
        refreshWorkbenchRowMaterial(fileId);
        return Promise.resolve();
      }
      var lastTs = _lastPutTsByFile[key] || 0;
      if (Date.now() - lastTs < 2500) {
        cachePatchCurrentRoleMap(fileId, nextMap);
        refreshWorkbenchRowMaterial(fileId);
        return Promise.resolve();
      }
    }
    cachePatchCurrentRoleMap(fileId, nextMap);
    if (String(selectedFileId) === String(fileId)) {
      currentRoleMap = nextMap;
    }
    refreshWorkbenchRowMaterial(fileId);
    return fetchJson(apiUrl('/api/sign/files/' + fileId + '/role-map'), {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ map: nextMap }),
    })
      .then(function (r) {
        var jj = (r && r.data) || {};
        if (jj.ok) {
          var finalMap = jj.map || nextMap;
          cachePatchCurrentRoleMap(fileId, finalMap);
          if (String(selectedFileId) === String(fileId)) {
            currentRoleMap = finalMap;
          }
          _lastPersistedRoleMapStable[key] = _roleMapStableJson(finalMap);
          _lastPutTsByFile[key] = Date.now();
        }
        updateSubmitState();
        // 不在此处 render：避免 render → scheduleAutoMatch → 再 PUT 的反馈环
      })
      .catch(function (e) {
        var msg = (e && e.message) || String(e || '保存角色映射失败');
        if (isBatchWorkbenchMode()) {
          setBatchWorkbenchMsg('角色映射保存失败：' + msg, true);
        }
      });
  }

  function ensureFileRoleLocales(fileId, row) {
    var st = fileUiCache[String(fileId)] || {};
    if (!st.roleLocales || typeof st.roleLocales !== 'object') {
      var base = row && row.locale === 'en' ? 'en' : row && row.locale === 'zh' ? 'zh' : 'auto';
      st.roleLocales = { author: base, reviewer: base, approver: base };
    }
    fileUiCache[String(fileId)] = st;
    return st.roleLocales;
  }

  function _signerNameForFilter(signer, hint) {
    if (signer && signer.name) return String(signer.name).trim();
    return String(hint || '').trim();
  }

  function _workbenchStrokeLocaleForRole(fileId, roleId, row) {
    var st = fileUiCache[String(fileId)] || {};
    var rl = st.roleLocales && st.roleLocales[roleId];
    if (rl === 'en' || rl === 'zh') return rl;
    if (row && row.locale === 'en') return 'en';
    if (row && row.locale === 'zh') return 'zh';
    var ctx = __aiwordHandoffCtxByFileId[String(fileId)] || {};
    return _isLikelyEnglishCountry(ctx.country) ? 'en' : 'zh';
  }

  function _strokeItemSignerId(kind, itemId) {
    var iid = String(itemId || '').trim();
    if (!iid) return '';
    var found = '';
    signersList.forEach(function (s) {
      if (found) return;
      var arr = kind === 'date' ? s.date_items || [] : s.sig_items || [];
      (arr || []).forEach(function (st) {
        if (String(st.id) === iid) found = String(s.id || '');
      });
    });
    return found;
  }

  function _strokeItemLocale(kind, itemId) {
    var iid = String(itemId || '').trim();
    if (!iid) return '';
    var loc = '';
    signersList.forEach(function (s) {
      if (loc) return;
      var arr = kind === 'date' ? s.date_items || [] : s.sig_items || [];
      (arr || []).forEach(function (st) {
        if (String(st.id) === iid) loc = st.locale === 'en' ? 'en' : 'zh';
      });
    });
    return loc;
  }

  function _appendStrokeOptionIfMissing(sel, kind, itemId) {
    var iid = String(itemId || '').trim();
    if (!iid) return;
    var has = Array.prototype.some.call(sel.options, function (op) {
      return op && op.value === iid;
    });
    if (has) return;
    signersList.forEach(function (s) {
      var arr = kind === 'date' ? s.date_items || [] : s.sig_items || [];
      (arr || []).forEach(function (st) {
        if (String(st.id) !== iid) return;
        var loc = st.locale === 'en' ? 'en' : 'zh';
        var o = document.createElement('option');
        o.value = st.id;
        o.setAttribute('data-signer-id', s.id);
        var tail = st.updated_at ? ' · ' + st.updated_at : '';
        o.textContent =
          (s.name || s.id) +
          ' · ' +
          (loc === 'en' ? '英文' : '中文') +
          ' · ' +
          (kind === 'date' ? '日期' : '签名') +
          ' · ' +
          (st.label || '') +
          tail;
        if (sel.options.length > 0) sel.insertBefore(o, sel.options[1] || null);
        else sel.appendChild(o);
      });
    });
  }

  function fillRoleItemSelectLocale(sel, kind, currentId, filterQ, strokeLocale) {
    sel.innerHTML = '';
    var o0 = document.createElement('option');
    o0.value = '';
    o0.textContent = kind === 'date' ? '请选择日期素材' : '请选择签名素材';
    sel.appendChild(o0);
    var q = filterQ ? String(filterQ).trim().toLowerCase() : '';
    var pl = strokeLocale === 'en' || strokeLocale === 'zh' ? strokeLocale : null;
    signersList.forEach(function (s) {
      if (q) {
        var nm = (s && s.name ? String(s.name) : '').toLowerCase();
        var sid = (s && s.id ? String(s.id) : '').toLowerCase();
        if (nm.indexOf(q) < 0 && sid.indexOf(q) < 0) return;
      }
      var arr = kind === 'date' ? s.date_items || [] : s.sig_items || [];
      (arr || []).forEach(function (st) {
        var loc = st.locale === 'en' ? 'en' : 'zh';
        if (pl && loc !== pl) return;
        var o = document.createElement('option');
        o.value = st.id;
        o.setAttribute('data-signer-id', s.id);
        var tail = st.updated_at ? ' · ' + st.updated_at : '';
        o.textContent =
          (s.name || s.id) +
          ' · ' +
          (loc === 'en' ? '英文' : '中文') +
          ' · ' +
          (kind === 'date' ? '日期' : '签名') +
          ' · ' +
          (st.label || '') +
          tail;
        sel.appendChild(o);
      });
    });
    _appendStrokeOptionIfMissing(sel, kind, currentId);
    if (currentId) {
      var ok = Array.prototype.some.call(sel.options, function (op) {
        return op.value === currentId;
      });
      if (!ok && pl) {
        // 已绑定素材的语言版本与当前「版本」筛选不一致时仍须展示，避免同一人跨文件看起来未匹配
        _appendStrokeOptionIfMissing(sel, kind, currentId);
        ok = Array.prototype.some.call(sel.options, function (op) {
          return op.value === currentId;
        });
      }
      if (ok) sel.value = currentId;
    }
  }

  function _ensureRoleFilterStateForFile(fileId, rid) {
    var st = fileUiCache[String(fileId)] || {};
    if (!st.roleItemFilterQ || typeof st.roleItemFilterQ !== 'object') {
      st.roleItemFilterQ = {};
    }
    if (!st.roleItemFilterQ[rid] || typeof st.roleItemFilterQ[rid] !== 'object') {
      st.roleItemFilterQ[rid] = { sig: '', date: '' };
    } else {
      if (typeof st.roleItemFilterQ[rid].sig !== 'string') st.roleItemFilterQ[rid].sig = '';
      if (typeof st.roleItemFilterQ[rid].date !== 'string') st.roleItemFilterQ[rid].date = '';
    }
    fileUiCache[String(fileId)] = st;
    return st.roleItemFilterQ[rid];
  }

  function syncWorkbenchFiltersFromMap(fileId, row) {
    var map = getFileRoleMapForWorkbench(fileId);
    ['author', 'reviewer', 'approver'].forEach(function (rid) {
      var p = map[rid];
      var fq = _ensureRoleFilterStateForFile(fileId, rid);
      var hint = _workbenchRoleNameHint(row, rid);
      if (hint) fq.sig = hint;
      if (p && p.sig) {
        var sid = _strokeItemSignerId('sig', p.sig);
        if (sid) {
          fq.date = _signerNameForFilter(_findSignerById(sid), hint);
        } else if (hint) {
          fq.date = hint;
        }
      } else if (hint) {
        fq.date = hint;
      }
    });
  }

  /**
   * 批量工作台：单文件「素材匹配」纯函数。
   * 输入：row 自身、当前已保存的 map（来自 fileUiCache）、签署人库快照。
   * 输出：{ map: 新 map, gaps: [缺口提示], changed: 是否需要 PUT }
   * 关键稳定性约束：本函数不发请求、不改全局；同一 row 多次调用结果一致（幂等）。
   */
  function _workbenchPlanRoleMap(fileId, row, existingMap) {
    var m = _deepCloneJsonish(existingMap || {});
    var gaps = [];
    var ctxIso = rowCtxFromRowState(row).doc_date;
    var handoffIso =
      (__aiwordHandoffCtxByFileId[String(fileId)] || {}).doc_date;
    var docIso = _parseAiwordDocDateIso(
      row.doc_date || ctxIso || handoffIso || ''
    );
    var signersReady = !!(signersList && signersList.length);

    // 仅 plan detect 命中过的角色：避免 executor 在普通文档误绑素材。
    var detectedSet = {};
    try {
      mergeDetectedRolesForFile(fileId).forEach(function (r) {
        if (r && r.id) detectedSet[String(r.id)] = true;
      });
    } catch (_) {}

    var rolesToPlan = ['author', 'reviewer', 'approver'];
    // executor 仅当 detect 已识别到（用例表/Tester 等强标签场景）才纳入 plan；
    // 其 hint 优先取 ctx.executor / row.executor，缺时回退到编写人员 → 复用 author 的素材。
    if (detectedSet['executor']) {
      rolesToPlan.push('executor');
    }

    rolesToPlan.forEach(function (rid) {
      var hint = _workbenchNameHintForRole(fileId, row, rid);
      var p = m[rid] && typeof m[rid] === 'object' ? Object.assign({}, m[rid]) : {};
      var roleNm = roleLabel(rid);
      var pl = _workbenchStrokeLocaleForRole(fileId, rid, row);

      // 步骤 1：按姓名匹配签署人；找不到 / 无姓名 时保留已有 p
      var signer = null;
      if (hint && signersReady) {
        signer = _findSignerByNameHint(hint);
        if (!signer) {
          // executor 找不到自身姓名签署人时，最后再尝试用 author 的签署人 / map.author 已绑定的 sig 来复用
          if (rid === 'executor') {
            var authorHint = _workbenchNameHintForRole(fileId, row, 'author');
            if (authorHint && authorHint !== hint) {
              var authorSigner = _findSignerByNameHint(authorHint);
              if (authorSigner) {
                signer = authorSigner;
                hint = authorHint;
              }
            }
          }
          if (!signer) {
            gaps.push(roleNm + '：库中未找到签署人「' + hint + '」');
          }
        }
      } else if (!hint && rid === 'executor' && signersReady) {
        // executor 无任何 hint：直接复用 author 已绑定的素材（map.author 或 author 的签署人）
        var authorMap = m['author'];
        if (authorMap && typeof authorMap === 'object' && authorMap.sig) {
          p.sig = p.sig || authorMap.sig;
          if (authorMap.date_mode && !p.date_mode) p.date_mode = authorMap.date_mode;
          if (authorMap.date_iso && !p.date_iso) p.date_iso = authorMap.date_iso;
          if (authorMap.date && !p.date) p.date = authorMap.date;
        } else {
          var aHint2 = _workbenchNameHintForRole(fileId, row, 'author');
          if (aHint2) {
            signer = _findSignerByNameHint(aHint2);
            if (signer) hint = aHint2;
          }
        }
      }

      // 步骤 2：签名素材（未绑定且库里有则补；已绑定则保留）
      if (signer && !p.sig) {
        var sigId =
          _firstStrokeId(signer, 'sig', pl) || _firstStrokeId(signer, 'sig', 'zh');
        if (sigId) p.sig = sigId;
      }

      // 步骤 3：日期素材
      //   - signersDbShare（共享库 + 拼接模式）：用 sig 实际语言定 date_mode；docIso 写入 date_iso
      //   - 非共享：尝试整张日期图（item 模式）
      if (signersDbShare) {
        var sigLoc = p.sig ? _strokeItemLocale('sig', p.sig) || pl : pl;
        var wantDm = sigLoc === 'en' ? 'composite_en_space' : 'composite_zh_ymd';
        if (!isCompositeDateMode(p.date_mode)) {
          p.date_mode = wantDm;
        } else if (p.sig && p.date_mode !== wantDm) {
          p.date_mode = wantDm;
        }
        // 拼接元件不全时勿写 date_iso（避免生成时整段日期失败）；优先整张日期图，否则仅签签名
        var sidForComp = p.sig ? _strokeItemSignerId('sig', p.sig) : '';
        var compMissing =
          docIso && sidForComp
            ? signerMissingCompositePieceKinds(sidForComp, wantDm, docIso)
            : [];
        if (compMissing.length) {
          var signerFb = sidForComp ? _findSignerById(sidForComp) : null;
          var dWhole =
            signerFb &&
            (_firstStrokeId(signerFb, 'date', pl) || _firstStrokeId(signerFb, 'date', 'zh'));
          if (dWhole) {
            p.date = dWhole;
            p.date_mode = 'item';
            p.date_iso = null;
          } else {
            p.date_iso = null;
            p.date = null;
            p.date_mode = wantDm;
          }
        } else if (docIso && p.date_iso !== docIso) {
          p.date_iso = docIso;
        }
        if (isCompositeDateMode(p.date_mode) && p.date) {
          p.date = null;
        }
      } else {
        if (signer && !p.date) {
          var dItemId =
            _firstStrokeId(signer, 'date', pl) || _firstStrokeId(signer, 'date', 'zh');
          if (dItemId) {
            p.date = dItemId;
            p.date_mode = 'item';
          }
        }
        if (!p.date_mode) p.date_mode = 'item';
      }

      // 步骤 4：写回；若 p 仍空则删除该角色（保持 map 紧凑）
      if (roleMapEntryNonEmpty(p)) {
        m[rid] = p;
      } else if (m[rid]) {
        delete m[rid];
      }

      // 步骤 5：缺口提示（不阻断，仅向用户解释将跳过的部分）
      if (hint && signer) {
        if (!p.sig) {
          gaps.push(roleNm + '（' + hint + '）：库中无该人签名素材，签名将跳过');
        }
        if (signersDbShare) {
          if (!p.date_iso && !p.date && p.sig) {
            gaps.push(roleNm + '（' + hint + '）：未填文档体现日期，日期将跳过');
          }
        } else if (!p.date && p.sig) {
          gaps.push(roleNm + '（' + hint + '）：库中无该人日期素材，日期将跳过');
        }
      } else if (hint && !signersReady) {
        gaps.push(roleNm + '（' + hint + '）：签署人库未就绪，请稍后点「批量匹配素材」重试');
      }
    });

    return { map: m, gaps: gaps, changed: _roleMapStableJson(m) !== _roleMapStableJson(existingMap || {}) };
  }

  // 同一 fileId 的 sync 同时只跑一个，避免多个入口（detect 完成、补日期、手改版本）互相覆盖
  var _workbenchSyncInflight = {};
  /** 已成功 PUT 的 map 快照（稳定 JSON），防止无操作下 plan 反复判定 changed 刷日志 */
  var _lastPersistedRoleMapStable = {};
  var __wbHandoffPipelineStarted = false;

  function workbenchSyncUiAfterMatch(fileId) {
    var key = String(fileId || '');
    if (!key) return Promise.resolve();
    var row = __batchWorkbenchRows[key];
    if (!row) return Promise.resolve();
    // 同一文件并发/重复调用只跑一轮，禁止排队导致整夜 PUT
    if (_workbenchSyncInflight[key]) {
      return _workbenchSyncInflight[key];
    }

    var p = (function () {
      var existing = getFileRoleMapForWorkbench(fileId);
      var plan = _workbenchPlanRoleMap(fileId, row, existing);
      var stablePlan = _roleMapStableJson(plan.map);
      if (plan.changed && _lastPersistedRoleMapStable[key] === stablePlan) {
        plan.changed = false;
      }
      var step;
      if (plan.changed) {
        cachePatchCurrentRoleMap(fileId, plan.map);
        if (String(selectedFileId) === String(fileId)) {
          currentRoleMap = _deepCloneJsonish(plan.map);
        }
        step = persistWorkbenchRoleMapChange(fileId, plan.map);
      } else {
        step = Promise.resolve();
      }
      return step
        .catch(function () {})
        .then(function () {
          syncWorkbenchFiltersFromMap(fileId, row);
          refreshWorkbenchRowMaterial(fileId);
          finalizeWorkbenchRowStatus(fileId);
          row._wbMatchAttempted = true;
          row._wbPipelineDone = true;
          renderBatchWorkbenchTable();
          if (plan.gaps && plan.gaps.length) {
            setBatchWorkbenchMsg(
              (row.name || fileId) + '：已载入可用素材；' + plan.gaps.join('；'),
              false
            );
          }
        });
    })();

    _workbenchSyncInflight[key] = p;
    p.finally(function () {
      if (_workbenchSyncInflight[key] === p) delete _workbenchSyncInflight[key];
    });
    return p;
  }

  // 兼容旧入口名：单点统一走 workbenchSyncUiAfterMatch
  function workbenchMatchRolesFromRowHints(fileId, row) {
    return workbenchSyncUiAfterMatch(fileId);
  }
  function workbenchApplyCompositeDefaults(fileId, row) {
    return workbenchSyncUiAfterMatch(fileId);
  }

  function _workbenchRoleNameHint(row, roleId) {
    if (!row) return '';
    if (roleId === 'author') return String(row.editor || '').trim();
    if (roleId === 'reviewer') return String(row.reviewer || '').trim();
    if (roleId === 'approver') return String(row.approver || '').trim();
    return '';
  }

  function patchWorkbenchRoleEntry(fileId, roleId, mutator) {
    var m = _deepCloneJsonish(getFileRoleMapForWorkbench(fileId));
    var p = m[roleId] && typeof m[roleId] === 'object' ? Object.assign({}, m[roleId]) : {};
    mutator(p, m);
    if (!roleMapEntryNonEmpty(p)) delete m[roleId];
    else m[roleId] = p;
    return persistWorkbenchRoleMapChange(fileId, m, { force: true }).then(function () {
      finalizeWorkbenchRowStatus(fileId);
      if (isBatchWorkbenchMode()) renderBatchWorkbenchTable();
    });
  }

  function workbenchRematchFileMaterials(fileId) {
    var row = __batchWorkbenchRows[String(fileId)];
    if (!row) return Promise.resolve();
    syncRowToHandoffCtx(fileId);
    return refreshSigners().then(function () {
      return workbenchSyncUiAfterMatch(fileId);
    });
  }

  function workbenchApplyDocDateToRoles(fileId) {
    return workbenchSyncUiAfterMatch(fileId);
  }

  function mkWorkbenchRoleSigCell(fileId, roleId, row) {
    var td = document.createElement('td');
    td.className = 'col-wb-sig';
    var stack = document.createElement('div');
    stack.className = 'batch-wb-cell-stack';
    var map = getFileRoleMapForWorkbench(fileId);
    var pair = map[roleId] && typeof map[roleId] === 'object' ? map[roleId] : {};
    var hint = _workbenchRoleNameHint(row, roleId);
    var strokeLoc = _workbenchStrokeLocaleForRole(fileId, roleId, row);
    var fq = _ensureRoleFilterStateForFile(fileId, roleId);
    if (hint && !fq.sig) fq.sig = hint;
    var filter = document.createElement('input');
    filter.type = 'search';
    filter.placeholder = '筛选';
    filter.value = fq.sig || hint || '';
    var sel = document.createElement('select');
    sel.title = '签名素材（' + (strokeLoc === 'en' ? '英文' : '中文') + '）';
    fillRoleItemSelectLocale(sel, 'sig', pair.sig || '', filter.value, strokeLoc);
    filter.addEventListener('click', function (ev) {
      ev.stopPropagation();
    });
    filter.addEventListener('input', function (ev) {
      ev.stopPropagation();
      fq.sig = filter.value || '';
      if (roleId === 'author') row.editor = fq.sig;
      else if (roleId === 'reviewer') row.reviewer = fq.sig;
      else if (roleId === 'approver') row.approver = fq.sig;
      syncRowToHandoffCtx(fileId);
      fillRoleItemSelectLocale(sel, 'sig', sel.value || '', fq.sig, strokeLoc);
    });
    sel.addEventListener('click', function (ev) {
      ev.stopPropagation();
    });
    sel.addEventListener('change', function (ev) {
      ev.stopPropagation();
      patchWorkbenchRoleEntry(fileId, roleId, function (p) {
        p.sig = sel.value || null;
      }).then(function () {
        try {
          var opt = sel.options[sel.selectedIndex];
          var sid2 = opt ? opt.getAttribute('data-signer-id') || '' : '';
          if (!sid2) return;
          var signer = _findSignerById(sid2);
          fq.date = _signerNameForFilter(signer, fq.sig || hint);
          var pl2 = _workbenchStrokeLocaleForRole(fileId, roleId, row);
          return patchWorkbenchRoleEntry(fileId, roleId, function (p2) {
            if (!p2.sig) return;
            var dItemId = _firstStrokeId(signer, 'date', pl2);
            if (dItemId) p2.date = dItemId;
            var iso = _parseAiwordDocDateIso((row && row.doc_date) || '');
            if (iso && signersDbShare) {
              p2.date_iso = iso;
              p2.date_mode = pl2 === 'en' ? 'composite_en_space' : 'composite_zh_ymd';
            }
          });
        } catch (_) {}
      });
    });
    stack.appendChild(filter);
    stack.appendChild(sel);
    td.appendChild(stack);
    return td;
  }

  function mkWorkbenchRoleDateCell(fileId, roleId, row) {
    var td = document.createElement('td');
    td.className = 'col-wb-date';
    var stack = document.createElement('div');
    stack.className = 'batch-wb-cell-stack';
    var map = getFileRoleMapForWorkbench(fileId);
    var pair = map[roleId] && typeof map[roleId] === 'object' ? map[roleId] : {};
    var hint = _workbenchRoleNameHint(row, roleId);
    var roleLocales = ensureFileRoleLocales(fileId, row);
    var dm0 = String(pair.date_mode || '').toLowerCase();
    if (dm0 === 'composite_en') dm0 = 'composite_en_space';
    if (signersDbShare) {
      if (['composite_zh_ymd', 'composite_en_space'].indexOf(dm0) < 0) {
        var plDef = _workbenchStrokeLocaleForRole(fileId, roleId, row);
        dm0 = plDef === 'en' ? 'composite_en_space' : 'composite_zh_ymd';
      }
    } else if (['composite_zh_ymd', 'composite_en_space'].indexOf(dm0) < 0) {
      dm0 = 'item';
    }
    var locLbl = document.createElement('label');
    locLbl.className = 'wb-mini-label';
    locLbl.textContent = '版本';
    var locSel = document.createElement('select');
    ;[
      { v: 'auto', t: '自动' },
      { v: 'zh', t: '中文' },
      { v: 'en', t: '英文' },
    ].forEach(function (o) {
      var opt = document.createElement('option');
      opt.value = o.v;
      opt.textContent = o.t;
      locSel.appendChild(opt);
    });
    locSel.value = roleLocales[roleId] || 'auto';
    locSel.addEventListener('click', function (ev) {
      ev.stopPropagation();
    });
    locSel.addEventListener('change', function (ev) {
      ev.stopPropagation();
      roleLocales[roleId] = locSel.value || 'auto';
      roleLocaleMap[roleId] = roleLocales[roleId];
      workbenchRematchFileMaterials(fileId);
    });
    var modeLbl = document.createElement('label');
    modeLbl.className = 'wb-mini-label';
    modeLbl.textContent = '日期方式';
    var dateModeSel = document.createElement('select');
    ;[
      { v: 'item', t: '整张日期图' },
      { v: 'composite_zh_ymd', t: '拼接·中文' },
      { v: 'composite_en_space', t: '拼接·英文' },
    ].forEach(function (o) {
      var opt = document.createElement('option');
      opt.value = o.v;
      opt.textContent = o.t;
      dateModeSel.appendChild(opt);
    });
    dateModeSel.value = dm0;
    if (!signersDbShare) dateModeSel.disabled = true;
    var itemBox = document.createElement('div');
    itemBox.className = 'batch-wb-date-item';
    var compBox = document.createElement('div');
    compBox.className = 'batch-wb-date-comp';
    var fq = _ensureRoleFilterStateForFile(fileId, roleId);
    if (hint && !fq.date) fq.date = hint;
    if (pair.sig && !fq.date) {
      var sid0 = _strokeItemSignerId('sig', pair.sig);
      if (sid0) {
        fq.date = _signerNameForFilter(_findSignerById(sid0), hint);
      }
    }
    var strokeLoc = _workbenchStrokeLocaleForRole(fileId, roleId, row);
    var dateFilter = document.createElement('input');
    dateFilter.type = 'search';
    dateFilter.placeholder = '筛选';
    dateFilter.value = fq.date || hint || '';
    var dateSel = document.createElement('select');
    fillRoleItemSelectLocale(dateSel, 'date', pair.date || '', dateFilter.value, strokeLoc);
    var dateIsoInp = document.createElement('input');
    dateIsoInp.type = 'date';
    dateIsoInp.value = _workbenchCompositeDateIso(row, pair);
    function applyDateModeUi(mode) {
      var comp = isCompositeDateMode(mode);
      itemBox.style.display = comp ? 'none' : 'flex';
      itemBox.style.flexDirection = 'column';
      itemBox.style.gap = '4px';
      compBox.style.display = comp ? 'flex' : 'none';
      compBox.style.flexDirection = 'column';
      compBox.style.gap = '4px';
    }
    dateFilter.addEventListener('click', function (ev) {
      ev.stopPropagation();
    });
    dateFilter.addEventListener('input', function (ev) {
      ev.stopPropagation();
      fq.date = dateFilter.value || '';
      fillRoleItemSelectLocale(dateSel, 'date', dateSel.value || '', fq.date, strokeLoc);
    });
    dateSel.addEventListener('click', function (ev) {
      ev.stopPropagation();
    });
    dateSel.addEventListener('change', function (ev) {
      ev.stopPropagation();
      patchWorkbenchRoleEntry(fileId, roleId, function (p) {
        p.date = dateSel.value || null;
        p.date_mode = 'item';
        p.date_iso = null;
      });
    });
    dateModeSel.addEventListener('change', function (ev) {
      ev.stopPropagation();
      var v = dateModeSel.value || 'item';
      patchWorkbenchRoleEntry(fileId, roleId, function (p) {
        if (v === 'item') {
          p.date_mode = 'item';
          p.date_iso = null;
        } else {
          p.date_mode = v;
          p.date = null;
          p.date_iso = dateIsoInp.value || _workbenchCompositeDateIso(row, p) || null;
        }
      }).then(function () {
        applyDateModeUi(v);
      });
    });
    dateIsoInp.addEventListener('click', function (ev) {
      ev.stopPropagation();
    });
    dateIsoInp.addEventListener('change', function (ev) {
      ev.stopPropagation();
      patchWorkbenchRoleEntry(fileId, roleId, function (p) {
        if (!isCompositeDateMode(p.date_mode)) return;
        p.date_iso = dateIsoInp.value || null;
      });
    });
    itemBox.appendChild(dateFilter);
    itemBox.appendChild(dateSel);
    compBox.appendChild(dateIsoInp);
    stack.appendChild(locLbl);
    stack.appendChild(locSel);
    if (signersDbShare) {
      stack.appendChild(modeLbl);
      stack.appendChild(dateModeSel);
    }
    stack.appendChild(itemBox);
    stack.appendChild(compBox);
    applyDateModeUi(dm0);
    td.appendChild(stack);
    return td;
  }

  function mkWorkbenchDocDateCell(fileId, row) {
    var td = document.createElement('td');
    td.className = 'col-wb-docdate';
    var inp = document.createElement('input');
    inp.type = 'date';
    inp.value = row.doc_date || '';
    inp.title = '文档体现日期（来自 aiword）';
    inp.addEventListener('click', function (ev) {
      ev.stopPropagation();
    });
    inp.addEventListener('change', function (ev) {
      ev.stopPropagation();
      row.doc_date = inp.value;
      syncRowToHandoffCtx(fileId);
      workbenchSyncUiAfterMatch(fileId);
    });
    td.appendChild(inp);
    return td;
  }

  function normalizeSavedFileRecords(files) {
    return (files || []).map(function (f) {
      if (!f) return f;
      var copy = Object.assign({}, f);
      if (copy.name) copy.name = repairDisplayFileName(copy.name);
      return copy;
    });
  }

  function _fileBaseNameOnly(name) {
    var s = String(name || '').trim();
    if (!s) return '';
    var i = Math.max(s.lastIndexOf('/'), s.lastIndexOf('\\'));
    return i >= 0 ? s.slice(i + 1) : s;
  }

  function fileHasValidDetectCache(fileId) {
    if (shouldSkipDetectPipelineForFile(fileId)) return true;
    var st = fileUiCache[String(fileId)] || {};
    var data = st.lastDetectData;
    return !!(data && data.ok && mergeDetectedRolesForFile(fileId).length);
  }

  function ensureFileRoleMapLoaded(fileId) {
    if (!fileId) return Promise.resolve();
    return fetchJson(apiUrl('/api/sign/files/' + fileId + '/role-map'), {
      timeoutMs: 120000,
    })
      .then(function (r) {
        var jj = (r && r.data) || {};
        if (jj.ok) {
          var loaded = jj.map || {};
          cachePatchCurrentRoleMap(fileId, loaded);
          _lastPersistedRoleMapStable[String(fileId)] = _roleMapStableJson(loaded);
          if (isBatchWorkbenchMode()) {
            finalizeWorkbenchRowStatus(fileId);
            try {
              renderBatchWorkbenchTable();
            } catch (_) {}
          }
        }
      })
      .catch(function () {});
  }

  function finalizeWorkbenchRowStatus(fileId) {
    var row = __batchWorkbenchRows[String(fileId)];
    if (!row) return;
    if (row.manualDetectWrong) {
      row.status = '识别有误';
      schedulePersistFileSessionCache(fileId);
      return;
    }
    syncWorkbenchRowFromDetect(fileId);
    if (isNoSignRequiredForFile(fileId)) {
      row.status = '无需签字';
      row.rolesLabel = '无需签字（用例表）';
      row.slotLabel = '无需签字';
      row.slotExplain = '规则标记该文档无需签字。';
      row.slotTags = ['无需签字'];
      row.slotSignable = true;
      row._wbPipelineDone = true;
      row._wbMatchAttempted = true;
      schedulePersistFileSessionCache(fileId);
      return;
    }
    refreshWorkbenchRowMaterial(fileId);
    var roles = mergeDetectedRolesForFile(fileId);
    var mat = row.material || assessFileMaterialStatus(fileId);
    if (!roles.length) {
      row.status = (fileUiCache[fileId] || {}).lastDetectError || '未识别到签字位';
      row.slotLabel = '不可签';
      row.slotBadgeClass = 'fail';
      row.slotRolesLine = '';
      row.slotLayoutLine = '';
      row.slotExplain = (fileUiCache[fileId] || {}).lastDetectError || '未识别到签字位';
      row.slotTags = ['未识别到签字位', '不可签'];
      row.slotSignable = false;
      return;
    }
    if (row.slotProbeOk === false) {
      row.status = '签字位不完整';
      if (!row.slotLabel || row.slotLabel === '—') {
        row.slotLabel = '不可签';
      }
      row.slotBadgeClass = row.slotBadgeClass || 'fail';
      row.slotTags = Array.isArray(row.slotTags) ? row.slotTags.slice() : [];
      if (row.slotTags.indexOf('不可签') < 0) row.slotTags.push('不可签');
      row.slotSignable = false;
      return;
    }
    if (row.slotMissingName || row.slotMissingDate) {
      // 结构分析显示缺位 → 主状态也应显示缺位，避免和签字位列「自相矛盾」
      var miss = [];
      if (row.slotMissingName) miss.push('姓名位');
      if (row.slotMissingDate) miss.push('日期位');
      row.status = '签字位不完整（缺' + miss.join('、') + '）';
      row.slotSignable = false;
      return;
    }
    if (mat.state === 'full' || (mat.matched > 0 && !mat.missingRoleIds.length)) {
      row.status = '就绪';
    } else if (mat.matched > 0) {
      row.status = '部分就绪';
    } else {
      row.status = '待匹配';
    }
    schedulePersistFileSessionCache(fileId);
  }

  function markWorkbenchRowDetectWrong(fileId, noteOpt, opts) {
    opts = opts || {};
    var row = __batchWorkbenchRows[String(fileId)];
    if (!row) return;
    row.status = '识别有误';
    row.manualDetectWrong = true;
    if (noteOpt) row.detectWrongNote = String(noteOpt);
    row._wbPipelineDone = true;
    if (!opts.skipPersist) schedulePersistFileSessionCache(fileId);
    if (!opts.skipRender) renderBatchWorkbenchTable();
  }

  var __fileDetectCorrectionCache = {};
  var __detectCorrectionEditingFileId = '';
  /** 批量登记时为目标文件 id 列表；单文件时长度为 1 */
  var __detectCorrectionEditingFileIds = [];
  var __detectCorrectionLoadGen = 0;
  var __detectCorrectionFormDirty = false;
  var __detectCorrectionFormDirtyBound = false;

  var DETECT_CORRECTION_ROLE_OPTIONS = [
    { id: 'author', label: '编写/编制' },
    { id: 'reviewer', label: '审核/复核' },
    { id: 'approver', label: '批准' },
    { id: 'executor', label: '执行/测试' },
    { id: 'reviewer_tail', label: '审核人员（文末）' },
  ];

  function _detectCorrectionModalEls() {
    return {
      modal: document.getElementById('detectCorrectionModal'),
      wrong: document.getElementById('detectCorrectionWrong'),
      saveRolesCb: document.getElementById('detectCorrectionSaveRolesCb'),
      saveSlotCb: document.getElementById('detectCorrectionSaveSlotCb'),
      rolesFieldset: document.getElementById('detectCorrectionRolesFieldset'),
      rolesBox: document.getElementById('detectCorrectionRoles'),
      slotFieldset: document.getElementById('detectCorrectionSlotFieldset'),
      slotArrangement: document.getElementById('detectCorrectionArrangement'),
      slotDateRelation: document.getElementById('detectCorrectionDateRelation'),
      slotDatePosition: document.getElementById('detectCorrectionDatePosition'),
      slotSeparator: document.getElementById('detectCorrectionSeparator'),
      note: document.getElementById('detectCorrectionNote'),
      keywords: document.getElementById('detectCorrectionKeywords'),
      imgInput: document.getElementById('detectCorrectionImageInput'),
      refList: document.getElementById('detectCorrectionRefList'),
      msg: document.getElementById('detectCorrectionModalMsg'),
      saveBtn: document.getElementById('detectCorrectionSaveBtn'),
      cancelBtn: document.getElementById('detectCorrectionCancelBtn'),
    };
  }

  function _collectExpectedSlotLayoutFromForm(els) {
    if (!els) return null;
    var out = {};
    if (els.slotArrangement && els.slotArrangement.value) {
      out.arrangement = String(els.slotArrangement.value);
    }
    if (els.slotDateRelation && els.slotDateRelation.value) {
      out.date_relation = String(els.slotDateRelation.value);
    }
    if (els.slotDatePosition && els.slotDatePosition.value) {
      out.date_position = String(els.slotDatePosition.value);
    }
    if (els.slotSeparator && els.slotSeparator.value) {
      out.separator = String(els.slotSeparator.value);
    }
    if (!Object.keys(out).length) return null;
    if (out.date_relation === 'same_cell') out.same_cell = true;
    if (out.date_relation === 'different_cell') out.same_cell = false;
    return out;
  }

  function _applyExpectedSlotLayoutToForm(els, esl) {
    esl = esl && typeof esl === 'object' ? esl : {};
    if (els.slotArrangement) {
      els.slotArrangement.value = String(esl.arrangement || '');
    }
    if (els.slotDateRelation) {
      els.slotDateRelation.value = String(esl.date_relation || '');
    }
    if (els.slotDatePosition) {
      els.slotDatePosition.value = String(esl.date_position || '');
    }
    if (els.slotSeparator) {
      els.slotSeparator.value = String(esl.separator || '');
    }
  }

  function _setDetectCorrectionModalMsg(text, isErr) {
    var els = _detectCorrectionModalEls();
    if (!els.msg) return;
    if (!text) {
      els.msg.style.display = 'none';
      els.msg.textContent = '';
      return;
    }
    els.msg.style.display = 'block';
    els.msg.style.color = isErr ? 'var(--error)' : 'var(--text-muted)';
    els.msg.textContent = text;
  }

  function _referenceImageUrl(fileId, imageId) {
    return (
      apiUrl(
        '/api/sign/files/' +
          encodeURIComponent(fileId) +
          '/detect-correction/reference-image/' +
          encodeURIComponent(imageId)
      ) + '?_=' + Date.now()
    );
  }

  function renderDetectCorrectionRefImages(fileId) {
    var els = _detectCorrectionModalEls();
    if (!els.refList) return;
    els.refList.innerHTML = '';
    var corr = __fileDetectCorrectionCache[String(fileId)] || {};
    var imgs = Array.isArray(corr.reference_images) ? corr.reference_images : [];
    imgs.forEach(function (im) {
      if (!im || !im.id) return;
      var wrap = document.createElement('div');
      wrap.className = 'sign-modal-ref-item';
      var img = document.createElement('img');
      img.src = _referenceImageUrl(fileId, im.id);
      img.alt = im.filename || '参考图';
      img.title = im.filename || '';
      var del = document.createElement('button');
      del.type = 'button';
      del.className = 'btn btn-secondary';
      del.textContent = '×';
      del.title = '删除';
      del.addEventListener('click', function () {
        fetchJson(
          apiUrl(
            '/api/sign/files/' +
              encodeURIComponent(fileId) +
              '/detect-correction/reference-image/' +
              encodeURIComponent(im.id)
          ),
          { method: 'DELETE', timeoutMs: 30000 }
        )
          .then(function (r) {
            var j = (r && r.data) || {};
            if (!j.ok) {
              _setDetectCorrectionModalMsg(j.error || '删除失败', true);
              return;
            }
            corr.reference_images = imgs.filter(function (x) {
              return String(x.id) !== String(im.id);
            });
            __fileDetectCorrectionCache[String(fileId)] = corr;
            renderDetectCorrectionRefImages(fileId);
          })
          .catch(function (e) {
            _setDetectCorrectionModalMsg((e && e.message) || '删除失败', true);
          });
      });
      wrap.appendChild(img);
      wrap.appendChild(del);
      els.refList.appendChild(wrap);
    });
  }

  function _bindDetectCorrectionFormDirtyOnce() {
    if (__detectCorrectionFormDirtyBound) return;
    var els = _detectCorrectionModalEls();
    if (!els.modal) return;
    __detectCorrectionFormDirtyBound = true;
    function markDirty() {
      __detectCorrectionFormDirty = true;
    }
    if (els.wrong) {
      els.wrong.addEventListener('input', markDirty);
      els.wrong.addEventListener('change', markDirty);
    }
    if (els.note) {
      els.note.addEventListener('input', markDirty);
      els.note.addEventListener('change', markDirty);
    }
    if (els.keywords) {
      els.keywords.addEventListener('input', markDirty);
      els.keywords.addEventListener('change', markDirty);
    }
    if (els.rolesBox) {
      els.rolesBox.addEventListener('change', markDirty);
    }
    [els.slotArrangement, els.slotDateRelation, els.slotDatePosition, els.slotSeparator].forEach(
      function (el) {
        if (!el) return;
        el.addEventListener('change', markDirty);
      }
    );
    [els.saveRolesCb, els.saveSlotCb].forEach(function (el) {
      if (!el) return;
      el.addEventListener('change', function () {
        markDirty();
        _applyDetectCorrectionSectionEnabled();
      });
    });
  }

  function _readDetectCorrectionSaveScopes(template) {
    template = template || {};
    var cs = template.correction_save;
    if (cs && typeof cs === 'object') {
      return {
        roles: cs.roles !== false,
        slot: cs.slot !== false,
      };
    }
    var hasRoles =
      Array.isArray(template.expected_roles) && template.expected_roles.length > 0;
    var esl = template.expected_slot_layout;
    var hasSlot =
      esl && typeof esl === 'object' && Object.keys(esl).length > 0;
    if (hasRoles && !hasSlot) return { roles: true, slot: false };
    if (hasSlot && !hasRoles) return { roles: false, slot: true };
    return { roles: true, slot: true };
  }

  function _applyDetectCorrectionSaveScopesToForm(template) {
    var els = _detectCorrectionModalEls();
    var scopes = _readDetectCorrectionSaveScopes(template);
    if (els.saveRolesCb) els.saveRolesCb.checked = scopes.roles;
    if (els.saveSlotCb) els.saveSlotCb.checked = scopes.slot;
    _applyDetectCorrectionSectionEnabled();
  }

  function _applyDetectCorrectionSectionEnabled() {
    var els = _detectCorrectionModalEls();
    var saveRoles = !els.saveRolesCb || els.saveRolesCb.checked;
    var saveSlot = !els.saveSlotCb || els.saveSlotCb.checked;
    if (els.rolesFieldset) {
      els.rolesFieldset.disabled = !saveRoles;
      els.rolesFieldset.style.opacity = saveRoles ? '1' : '0.55';
    }
    if (els.slotFieldset) {
      els.slotFieldset.disabled = !saveSlot;
      els.slotFieldset.style.opacity = saveSlot ? '1' : '0.55';
    }
  }

  function _applyDetectCorrectionFormFromTemplate(els, template, detectedRoleIds) {
    if (__detectCorrectionFormDirty) return;
    template = template || {};
    if (els.wrong) els.wrong.value = String(template.wrong_description || '');
    if (els.note) els.note.value = String(template.expected_note || '');
    if (els.keywords) {
      var kws = Array.isArray(template.label_keywords) ? template.label_keywords : [];
      els.keywords.value = kws.join('，');
    }
    var sel = {};
    var exp = Array.isArray(template.expected_roles) ? template.expected_roles : [];
    if (exp.length) {
      exp.forEach(function (rid) {
        sel[rid] = true;
      });
    } else {
      (detectedRoleIds || []).forEach(function (rid) {
        sel[rid] = true;
      });
    }
    _buildDetectCorrectionRoleChecks(sel);
    _applyExpectedSlotLayoutToForm(els, template.expected_slot_layout || {});
    _applyDetectCorrectionSaveScopesToForm(template);
  }

  function _buildDetectCorrectionRoleChecks(selectedMap) {
    var els = _detectCorrectionModalEls();
    if (!els.rolesBox) return;
    els.rolesBox.innerHTML = '';
    DETECT_CORRECTION_ROLE_OPTIONS.forEach(function (opt) {
      var lab = document.createElement('label');
      var cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.value = opt.id;
      cb.checked = !!(selectedMap && selectedMap[opt.id]);
      lab.appendChild(cb);
      lab.appendChild(document.createTextNode(opt.label || roleLabel(opt.id)));
      els.rolesBox.appendChild(lab);
    });
  }

  function loadDetectCorrectionForFile(fileId) {
    if (!fileId) return Promise.resolve({});
    if (__fileDetectCorrectionCache[String(fileId)]) {
      return Promise.resolve(__fileDetectCorrectionCache[String(fileId)]);
    }
    return fetchJson(apiUrl('/api/sign/files/' + encodeURIComponent(fileId) + '/detect-correction'), {
      timeoutMs: 30000,
    })
      .then(function (r) {
        var j = (r && r.data) || {};
        var corr = j.ok && j.correction ? j.correction : {};
        __fileDetectCorrectionCache[String(fileId)] = corr;
        return corr;
      })
      .catch(function () {
        return {};
      });
  }

  function _getWorkbenchSelectedFileIds() {
    var ids = [];
    savedFiles.forEach(function (rec) {
      if (!rec || !rec.id) return;
      var row = __batchWorkbenchRows[String(rec.id)];
      if (row && row.selected) ids.push(String(rec.id));
    });
    return ids;
  }

  function _renderDetectCorrectionBatchFileList(fileIds) {
    var ul = document.getElementById('detectCorrectionBatchFileList');
    var scopeHint = document.getElementById('detectCorrectionScopeHint');
    if (!ul) return;
    ul.innerHTML = '';
    if (!fileIds || fileIds.length <= 1) {
      ul.style.display = 'none';
      if (scopeHint) {
        scopeHint.textContent =
          '说明错在哪；勾选本次要保存的范围。保存后点「识别」会重跑需签角色与签字位并带入登记。';
      }
      return;
    }
    ul.style.display = 'block';
    if (scopeHint) {
      scopeHint.textContent =
        '以下登记将应用到已勾选的 ' +
        fileIds.length +
        ' 个文件（各文件会记录当时识别到的错误角色）。参考图上传至 FTP 后，各文件登记共用同一 ftp_path。';
    }
    fileIds.forEach(function (fid) {
      var rec = savedFiles.find(function (x) {
        return x && String(x.id) === String(fid);
      });
      var li = document.createElement('li');
      li.textContent = rec && rec.name ? _fileBaseNameOnly(rec.name) : fid;
      ul.appendChild(li);
    });
  }

  function _updateDetectCorrectionSaveBtnLabel(n) {
    var btn = document.getElementById('detectCorrectionSaveBtn');
    if (!btn) return;
    if (n > 1) {
      btn.textContent = '保存并标为识别有误（' + n + ' 个文件）';
    } else {
      btn.textContent = '保存并标为识别有误';
    }
  }

  function openDetectCorrectionDialog(fileIdOrIds, opt) {
    opt = opt && typeof opt === 'object' ? opt : {};
    var els = _detectCorrectionModalEls();
    var ids = [];
    if (Array.isArray(fileIdOrIds)) {
      ids = fileIdOrIds.map(String).filter(Boolean);
    } else if (fileIdOrIds) {
      ids = [String(fileIdOrIds)];
    }
    if (!ids.length) return;
    if (!els.modal) {
      setBatchWorkbenchMsg('登记表单未加载，请 Ctrl+F5 强刷页面后重试', true);
      return;
    }
    __detectCorrectionEditingFileIds = ids.slice();
    __detectCorrectionEditingFileId = ids[0];
    __detectCorrectionLoadGen += 1;
    var loadGen = __detectCorrectionLoadGen;
    __detectCorrectionFormDirty = false;
    _bindDetectCorrectionFormDirtyOnce();
    _setDetectCorrectionModalMsg('', false);
    _updateDetectCorrectionSaveBtnLabel(ids.length);
    _renderDetectCorrectionBatchFileList(ids);
    var title = document.getElementById('detectCorrectionTitle');
    if (title) {
      if (ids.length > 1) {
        title.textContent = '批量登记识别纠正（' + ids.length + ' 个文件）';
      } else {
        var rec0 = savedFiles.find(function (x) {
          return x && String(x.id) === ids[0];
        });
        title.textContent =
          '登记识别纠正' + (rec0 && rec0.name ? ' — ' + _fileBaseNameOnly(rec0.name) : '');
      }
    }
    var scopeHint = document.getElementById('detectCorrectionScopeHint');
    if (scopeHint && ids.length <= 1) {
      scopeHint.textContent =
        '说明错在哪；勾选本次要保存的范围（需签角色 / 签字位版式可只选其一）。保存后点「识别」会同时重跑角色与签字位并带入登记。';
    }
    if (els.saveRolesCb) els.saveRolesCb.checked = true;
    if (els.saveSlotCb) els.saveSlotCb.checked = true;
    _applyDetectCorrectionSectionEnabled();
    if (els.wrong) {
      els.wrong.value = '';
      els.wrong.placeholder =
        '例如：把用例表列头「测试人」误识别为编写；或漏识别批准人；或日期应在角色下方而非右方';
    }
    if (els.note) els.note.value = '';
    if (els.keywords) els.keywords.value = '';
    _applyExpectedSlotLayoutToForm(els, {});
    _buildDetectCorrectionRoleChecks({});
    renderDetectCorrectionRefImages(ids[0]);
    els.modal.classList.add('show');
    els.modal.setAttribute('aria-hidden', 'false');
    var primaryId = ids[0];
    var detected = mergeDetectedRolesForFile(primaryId).map(function (x) {
      return x.id;
    });
    if (ids.length > 1) {
      _setDetectCorrectionModalMsg('正在加载已有登记（仅首条）…', false);
    }
    loadDetectCorrectionForFile(primaryId)
      .then(function (template) {
        if (loadGen !== __detectCorrectionLoadGen) return;
        _applyDetectCorrectionFormFromTemplate(els, template || {}, detected);
        renderDetectCorrectionRefImages(primaryId);
        if (ids.length > 1) {
          _setDetectCorrectionModalMsg('', false);
        }
      })
      .catch(function (e) {
        if (loadGen !== __detectCorrectionLoadGen) return;
        _setDetectCorrectionModalMsg(
          '加载历史登记失败：' + ((e && e.message) || '请检查服务是否已重启'),
          true
        );
      });
  }

  function closeDetectCorrectionDialog() {
    var els = _detectCorrectionModalEls();
    if (!els.modal) return;
    els.modal.classList.remove('show');
    els.modal.setAttribute('aria-hidden', 'true');
    __detectCorrectionEditingFileId = '';
    __detectCorrectionEditingFileIds = [];
    _setDetectCorrectionModalMsg('', false);
    _renderDetectCorrectionBatchFileList([]);
    _updateDetectCorrectionSaveBtnLabel(1);
  }

  function _collectDetectCorrectionSharedFromForm() {
    var els = _detectCorrectionModalEls();
    var wrong = els.wrong ? String(els.wrong.value || '').trim() : '';
    if (!wrong) return { error: '请填写「错在哪」' };
    var saveRoles = !els.saveRolesCb || els.saveRolesCb.checked;
    var saveSlot = !els.saveSlotCb || els.saveSlotCb.checked;
    if (!saveRoles && !saveSlot) {
      return { error: '请至少勾选一项「本次保存范围」' };
    }
    var expected = [];
    if (els.rolesBox) {
      var cbs = els.rolesBox.querySelectorAll('input[type="checkbox"]');
      for (var i = 0; i < cbs.length; i++) {
        if (cbs[i].checked) expected.push(String(cbs[i].value || '').trim());
      }
    }
    var kws = [];
    if (els.keywords) {
      String(els.keywords.value || '')
        .split(/[,，;；\s]+/)
        .forEach(function (p) {
          p = String(p || '').trim();
          if (p && kws.indexOf(p) < 0) kws.push(p);
        });
    }
    var slotLayout = saveSlot ? _collectExpectedSlotLayoutFromForm(els) : null;
    return {
      shared: {
        wrong_description: wrong,
        expected_roles: expected,
        expected_note: els.note ? String(els.note.value || '').trim() : '',
        label_keywords: kws,
        expected_slot_layout: slotLayout,
        saveRoles: saveRoles,
        saveSlot: saveSlot,
        correction_save: { roles: saveRoles, slot: saveSlot },
      },
    };
  }

  function _collectDetectCorrectionPayloadForFile(fileId, shared) {
    var prev = __fileDetectCorrectionCache[String(fileId)] || {};
    var payload = {
      wrong_description: shared.wrong_description,
      expected_note: shared.expected_note,
      wrong_roles_detected: mergeDetectedRolesForFile(fileId).map(function (x) {
        return x.id;
      }),
      label_keywords: shared.label_keywords.slice(),
      reference_images: Array.isArray(prev.reference_images) ? prev.reference_images.slice() : [],
      correction_save: shared.correction_save || { roles: true, slot: true },
    };
    if (shared.saveRoles) {
      payload.expected_roles = shared.expected_roles.slice();
    } else if (Array.isArray(prev.expected_roles)) {
      payload.expected_roles = prev.expected_roles.slice();
    } else {
      payload.expected_roles = [];
    }
    if (shared.saveSlot) {
      if (shared.expected_slot_layout) {
        payload.expected_slot_layout = shared.expected_slot_layout;
      }
    } else if (prev.expected_slot_layout) {
      payload.expected_slot_layout = prev.expected_slot_layout;
    }
    return payload;
  }

  function saveDetectCorrectionDialog() {
    var ids =
      __detectCorrectionEditingFileIds && __detectCorrectionEditingFileIds.length
        ? __detectCorrectionEditingFileIds.slice()
        : __detectCorrectionEditingFileId
          ? [__detectCorrectionEditingFileId]
          : [];
    if (!ids.length) return Promise.resolve();
    var built = _collectDetectCorrectionSharedFromForm();
    if (built.error) {
      _setDetectCorrectionModalMsg(built.error, true);
      return Promise.resolve();
    }
    var els = _detectCorrectionModalEls();
    if (els.saveBtn) els.saveBtn.disabled = true;
    _setDetectCorrectionModalMsg('正在保存…', false);
    var noteShort = (built.shared.wrong_description || '').slice(0, 80);
    var items = ids.map(function (fid) {
      return {
        file_id: fid,
        correction: _collectDetectCorrectionPayloadForFile(fid, built.shared),
      };
    });

    function _afterSaveResults(okIds, failN, saveResp) {
      var j = saveResp || {};
      (okIds || []).forEach(function (fid) {
        var payload = null;
        items.forEach(function (it) {
          if (String(it.file_id) === String(fid)) payload = it.correction;
        });
        if (payload) {
          __fileDetectCorrectionCache[String(fid)] = payload;
        }
        markWorkbenchRowDetectWrong(fid, '已登记纠正：' + noteShort, {
          skipPersist: true,
          skipRender: true,
        });
      });
      closeDetectCorrectionDialog();
      try {
        renderBatchWorkbenchTable();
      } catch (_) {}
      var okN = (okIds || []).length;
      var ruleHint = '';
      if (ids.length > 1 && Array.isArray(j.rule_syncs)) {
        var rsN = 0;
        j.rule_syncs.forEach(function (rs) {
          if (rs && rs.ok) rsN++;
        });
        if (rsN) ruleHint = '；已同步角色/签字位规则至 JSON（' + rsN + ' 条）';
      } else if (j.rule_sync) {
        var oneHint = _formatRuleSyncHint(j.rule_sync);
        if (oneHint) ruleHint = '；' + oneHint;
      }
      if (ids.length > 1) {
        setBatchWorkbenchMsg(
          (failN
            ? '批量登记完成：成功 ' + okN + ' 个，失败 ' + failN + ' 个'
            : '已对 ' + okN + ' 个文件保存纠正登记，重新识别时将自动带入') + ruleHint,
          !!failN
        );
      } else if (okN) {
        setBatchWorkbenchMsg(
          '已保存识别纠正登记，重新识别时将自动带入' + ruleHint,
          false
        );
      } else {
        setBatchWorkbenchMsg('保存失败', true);
      }
      flushLightFileSessionCaches(okIds).catch(function () {});
    }

    if (ids.length > 1) {
      var batchTimeout = _detectCorrectionBatchTimeoutMs(ids.length);
      _setDetectCorrectionModalMsg(
        '正在批量保存 ' + ids.length + ' 个文件（最长约 ' + Math.round(batchTimeout / 1000) + 's）…',
        false
      );
      return fetchJsonWithRetry(
        apiUrl('/api/sign/files/detect-correction/batch'),
        {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ items: items }),
          timeoutMs: batchTimeout,
        },
        {
          maxTry: 2,
          delayMs: 1500,
          onRetry: function (n) {
            _setDetectCorrectionModalMsg('保存较慢，正在重试（' + n + '/2）…', false);
          },
        }
      )
        .then(function (r) {
          var j = (r && r.data) || {};
          if (!j.ok) {
            _setDetectCorrectionModalMsg(j.error || '批量保存失败', true);
            return;
          }
          var okIds = Array.isArray(j.ok_ids) ? j.ok_ids : [];
          var failN = Number(j.fail_count) || 0;
          if (failN && Array.isArray(j.failed) && j.failed.length) {
            console.warn('[detect-correction batch]', j.failed);
          }
          return _afterSaveResults(okIds, failN, j);
        })
        .catch(function (e) {
          _setDetectCorrectionModalMsg((e && e.message) || '批量保存失败', true);
        })
        .finally(function () {
          if (els.saveBtn) els.saveBtn.disabled = false;
        });
    }

    var fid0 = ids[0];
    var payload0 = items[0].correction;
    return fetchJson(apiUrl('/api/sign/files/' + encodeURIComponent(fid0) + '/detect-correction'), {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ correction: payload0 }),
      timeoutMs: 60000,
    })
      .then(function (r) {
        var j = (r && r.data) || {};
        if (!j.ok) {
          _setDetectCorrectionModalMsg(j.error || '保存失败', true);
          return;
        }
        __fileDetectCorrectionCache[String(fid0)] = j.correction || payload0;
        return _afterSaveResults([fid0], 0, j);
      })
      .catch(function (e) {
        _setDetectCorrectionModalMsg((e && e.message) || '保存失败', true);
      })
      .finally(function () {
        if (els.saveBtn) els.saveBtn.disabled = false;
      });
  }

  function _attachDetectCorrectionImageToFiles(ids, imageMeta) {
    if (!imageMeta || !imageMeta.id) return;
    ids.forEach(function (fid) {
      var corr = __fileDetectCorrectionCache[String(fid)] || {};
      var imgs = Array.isArray(corr.reference_images) ? corr.reference_images.slice() : [];
      var dup = imgs.some(function (x) {
        return (
          String(x.id) === String(imageMeta.id) ||
          (imageMeta.ftp_path && String(x.ftp_path) === String(imageMeta.ftp_path))
        );
      });
      if (!dup) imgs.push(imageMeta);
      corr.reference_images = imgs;
      __fileDetectCorrectionCache[String(fid)] = corr;
    });
  }

  function uploadDetectCorrectionImages(fileIdOrIds, fileList) {
    var ids = [];
    if (Array.isArray(fileIdOrIds)) {
      ids = fileIdOrIds.map(String).filter(Boolean);
    } else if (fileIdOrIds) {
      ids = [String(fileIdOrIds)];
    }
    if (!ids.length || !fileList || !fileList.length) return Promise.resolve();
    var chain = Promise.resolve();
    Array.prototype.forEach.call(fileList, function (f) {
      chain = chain.then(function () {
        var fd = new FormData();
        fd.append('file', f);
        if (ids.length > 1) {
          return fetch(apiUrl('/api/sign/detect-correction/reference-image'), {
            method: 'POST',
            body: fd,
            credentials: 'same-origin',
          }).then(function (resp) {
            return resp.json().then(function (j) {
              if (!resp.ok || !j.ok) throw new Error(j.error || '上传失败');
              _attachDetectCorrectionImageToFiles(ids, j.image);
            });
          });
        }
        var fid = ids[0];
        return fetch(
          apiUrl(
            '/api/sign/files/' +
              encodeURIComponent(fid) +
              '/detect-correction/reference-image'
          ),
          { method: 'POST', body: fd, credentials: 'same-origin' }
        ).then(function (resp) {
          return resp.json().then(function (j) {
            if (!resp.ok || !j.ok) throw new Error(j.error || '上传失败');
            _attachDetectCorrectionImageToFiles([fid], j.image);
          });
        });
      });
    });
    return chain.then(function () {
      renderDetectCorrectionRefImages(ids[0]);
      if (ids.length > 1) {
        _setDetectCorrectionModalMsg(
          '参考图已上传至 FTP，' + ids.length + ' 个文件共用同一文件链接',
          false
        );
      }
    });
  }

  function initDetectCorrectionModal() {
    var els = _detectCorrectionModalEls();
    if (!els.modal) return;
    if (els.cancelBtn) {
      els.cancelBtn.addEventListener('click', function () {
        closeDetectCorrectionDialog();
      });
    }
    if (els.saveBtn) {
      els.saveBtn.addEventListener('click', function () {
        saveDetectCorrectionDialog();
      });
    }
    if (els.imgInput) {
      els.imgInput.addEventListener('change', function () {
        var fids =
          __detectCorrectionEditingFileIds && __detectCorrectionEditingFileIds.length
            ? __detectCorrectionEditingFileIds.slice()
            : __detectCorrectionEditingFileId
              ? [__detectCorrectionEditingFileId]
              : [];
        if (!fids.length || !els.imgInput.files || !els.imgInput.files.length) return;
        _setDetectCorrectionModalMsg('正在上传参考图…', false);
        uploadDetectCorrectionImages(fids, els.imgInput.files)
          .then(function () {
            if (fids.length <= 1) {
              _setDetectCorrectionModalMsg('参考图已上传', false);
            }
            els.imgInput.value = '';
          })
          .catch(function (e) {
            _setDetectCorrectionModalMsg((e && e.message) || '上传失败', true);
          });
      });
    }
    els.modal.addEventListener('click', function (ev) {
      if (ev.target === els.modal) closeDetectCorrectionDialog();
    });
  }

  function markWorkbenchSelectedAsDetectWrong() {
    var selected = _getWorkbenchSelectedFileIds();
    if (!selected.length) {
      setBatchWorkbenchMsg('请先勾选要批量登记的文件', true);
      return;
    }
    openDetectCorrectionDialog(selected);
  }

  function _workbenchStatusBucket(shown) {
    var t = String(shown || '').trim();
    if (t === '识别有误') return '识别有误';
    if (t === '待匹配') return '待匹配';
    if (t === '无可用素材') return '无可用素材';
    if (/识别失败/.test(t)) return '识别失败';
    if (/识别超时/.test(t)) return '识别超时';
    if (/未识别/.test(t)) return '未识别';
    if (/签字位不完整/.test(t)) return '签字位不完整';
    if (t === '就绪') return '就绪';
    if (t === '部分就绪') return '部分就绪';
    if (t === '无需签字') return '无需签字';
    if (/^已签字/.test(t) || t === '已签字') return '已签字';
    if (/签字失败/.test(t)) return '签字失败';
    if (/保存映射失败/.test(t)) return '保存映射失败';
    if (/^签字中/.test(t)) return '签字中';
    if (/识别|匹配|分析/.test(t)) return '识别中';
    if (t === '—' || !t || t === '待处理') return '待处理';
    return t;
  }

  function _getWorkbenchSlotFilterCatalogMap() {
    var map = {};
    WORKBENCH_SLOT_FILTER_GROUPS.forEach(function (grp) {
      (grp.items || []).forEach(function (it) {
        var v = _normalizeWorkbenchSlotTag(it && it.value != null ? it.value : it);
        if (!v) return;
        map[v] = it.label || v;
      });
    });
    return map;
  }

  function _workbenchSlotFilterGroupForTag(tag) {
    var t = _normalizeWorkbenchSlotTag(tag);
    if (!t) return '其它';
    var i;
    for (i = 0; i < WORKBENCH_SLOT_FILTER_GROUPS.length; i++) {
      var grp = WORKBENCH_SLOT_FILTER_GROUPS[i];
      var found = (grp.items || []).some(function (it) {
        var v = _normalizeWorkbenchSlotTag(it && it.value != null ? it.value : it);
        return v === t;
      });
      if (found) return grp.title;
    }
    if (/^版式-/.test(t)) return '版式';
    if (/^分隔-/.test(t)) return '分隔方式';
    if (/识别|未识别/.test(t)) return '识别过程';
    return '其它';
  }

  function _buildWorkbenchSlotFilterGroupsForUi(extraTags) {
    var catalogMap = _getWorkbenchSlotFilterCatalogMap();
    var out = WORKBENCH_SLOT_FILTER_GROUPS.map(function (grp) {
      return {
        title: grp.title,
        items: (grp.items || []).map(function (it) {
          var value = _normalizeWorkbenchSlotTag(it && it.value != null ? it.value : it);
          return { value: value, label: it.label || value };
        }),
      };
    });
    var extras = [];
    (extraTags || []).forEach(function (tag) {
      var v = _normalizeWorkbenchSlotTag(tag);
      if (!v || catalogMap[v]) return;
      if (extras.indexOf(v) < 0) extras.push(v);
    });
    extras.sort();
    if (extras.length) {
      out.push({
        title: '其它（本次列表）',
        items: extras.map(function (v) {
          return { value: v, label: v };
        }),
      });
    }
    return out;
  }

  function _workbenchRowSlotFilterKeys(row) {
    var keys = [];
    if (!row) return keys;
    function add(k) {
      k = _normalizeWorkbenchSlotTag(k);
      if (!k || keys.indexOf(k) >= 0) return;
      keys.push(k);
    }
    var tags = Array.isArray(row.slotTags) ? row.slotTags : [];
    tags.forEach(add);
    add(row.slotLabel);
    return keys;
  }

  function _slotFilterValueMatchesRow(filterValue, rowKeys) {
    filterValue = _normalizeWorkbenchSlotTag(filterValue);
    if (!filterValue) return false;
    var aliases = WORKBENCH_SLOT_FILTER_MATCH_KEYS[filterValue];
    var candidates = aliases ? aliases.slice() : [filterValue];
    if (candidates.indexOf(filterValue) < 0) candidates.unshift(filterValue);
    for (var i = 0; i < candidates.length; i++) {
      if (rowKeys.indexOf(candidates[i]) >= 0) return true;
    }
    return false;
  }

  function _renderGroupedWorkbenchFilterBox(boxEl, groups, dataAttr, onChange) {
    if (!boxEl) return;
    boxEl.innerHTML = '';
    (groups || []).forEach(function (grp) {
      var items = grp.items || [];
      if (!items.length) return;
      var section = document.createElement('div');
      section.className = 'wb-filter-group';
      var titleEl = document.createElement('span');
      titleEl.className = 'wb-filter-group-title';
      titleEl.textContent = grp.title || '';
      section.appendChild(titleEl);
      var itemsWrap = document.createElement('div');
      itemsWrap.className = 'wb-filter-group-items';
      items.forEach(function (item) {
        var value = typeof item === 'string' ? item : item.value;
        var label = typeof item === 'string' ? item : item.label || item.value;
        if (!value) return;
        var lab = document.createElement('label');
        var cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.setAttribute(dataAttr, value);
        if (item && item.checked) cb.checked = true;
        if (onChange) {
          cb.addEventListener('change', onChange);
        }
        lab.appendChild(cb);
        lab.appendChild(document.createTextNode(label));
        itemsWrap.appendChild(lab);
      });
      section.appendChild(itemsWrap);
      boxEl.appendChild(section);
    });
  }

  function _initWorkbenchStatusFilterCheckboxes() {
    if (!batchWorkbenchFilterStatusBox) return;
    _renderGroupedWorkbenchFilterBox(
      batchWorkbenchFilterStatusBox,
      WORKBENCH_STATUS_FILTER_GROUPS,
      'data-wb-status'
    );
  }

  function _workbenchStatusFilterIsAllSelected() {
    if (!batchWorkbenchFilterStatusBox) return false;
    var cbs = batchWorkbenchFilterStatusBox.querySelectorAll('input[data-wb-status]');
    if (!cbs.length) return false;
    var n = 0;
    for (var i = 0; i < cbs.length; i++) {
      if (cbs[i].checked) n++;
    }
    return n >= cbs.length;
  }

  function _updateWorkbenchStatusFilterToggleLabel() {
    if (!batchWorkbenchFilterStatusToggle) return;
    var n = __wbFilterStatuses.length;
    if (!n) {
      batchWorkbenchFilterStatusToggle.textContent = '状态筛选 ▼';
      return;
    }
    if (_workbenchStatusFilterIsAllSelected()) {
      batchWorkbenchFilterStatusToggle.textContent = '状态：全部 ▼';
      return;
    }
    batchWorkbenchFilterStatusToggle.textContent = '状态：已选 ' + n + ' 项 ▼';
  }

  function _isWorkbenchDetectIssueStatus(shown) {
    var b = _workbenchStatusBucket(shown);
    return (
      b === '识别有误' ||
      b === '识别失败' ||
      b === '识别超时' ||
      b === '未识别'
    );
  }

  function _readWorkbenchFilterStatusesFromUi() {
    if (!batchWorkbenchFilterStatusBox) return [];
    var out = [];
    var cbs = batchWorkbenchFilterStatusBox.querySelectorAll('input[data-wb-status]');
    for (var i = 0; i < cbs.length; i++) {
      if (cbs[i].checked) {
        var v = String(cbs[i].getAttribute('data-wb-status') || '').trim();
        if (v) out.push(v);
      }
    }
    return out;
  }

  function _bindWorkbenchStatusFilterCheckboxes() {
    _initWorkbenchStatusFilterCheckboxes();
    if (!batchWorkbenchFilterStatusBox) return;
    var cbs = batchWorkbenchFilterStatusBox.querySelectorAll('input[data-wb-status]');
    for (var i = 0; i < cbs.length; i++) {
      cbs[i].addEventListener('change', function () {
        __wbFilterStatuses = _readWorkbenchFilterStatusesFromUi();
        _updateWorkbenchStatusFilterToggleLabel();
        renderBatchWorkbenchTable();
      });
    }
  }

  function _setWorkbenchFilterStatusChecks(statusKeys) {
    if (!batchWorkbenchFilterStatusBox) return;
    var keys = statusKeys || {};
    var cbs = batchWorkbenchFilterStatusBox.querySelectorAll('input[data-wb-status]');
    for (var i = 0; i < cbs.length; i++) {
      var v = String(cbs[i].getAttribute('data-wb-status') || '');
      cbs[i].checked = !!keys[v];
    }
    __wbFilterStatuses = _readWorkbenchFilterStatusesFromUi();
    _updateWorkbenchStatusFilterToggleLabel();
  }

  function _selectAllWorkbenchStatusFilters() {
    if (!batchWorkbenchFilterStatusBox) return;
    var cbs = batchWorkbenchFilterStatusBox.querySelectorAll('input[data-wb-status]');
    for (var i = 0; i < cbs.length; i++) cbs[i].checked = true;
    __wbFilterStatuses = _readWorkbenchFilterStatusesFromUi();
    _updateWorkbenchStatusFilterToggleLabel();
    renderBatchWorkbenchTable();
  }

  function _clearWorkbenchStatusFilters() {
    if (!batchWorkbenchFilterStatusBox) return;
    var cbs = batchWorkbenchFilterStatusBox.querySelectorAll('input[data-wb-status]');
    for (var i = 0; i < cbs.length; i++) cbs[i].checked = false;
    __wbFilterStatuses = [];
    _updateWorkbenchStatusFilterToggleLabel();
    renderBatchWorkbenchTable();
  }

  function _toggleWorkbenchStatusFilterPanel(open) {
    if (!batchWorkbenchFilterStatusWrap) return;
    if (open) batchWorkbenchFilterStatusWrap.classList.add('open');
    else batchWorkbenchFilterStatusWrap.classList.remove('open');
  }

  function _normalizeWorkbenchSlotTag(tag) {
    return String(tag || '').trim();
  }

  function _collectWorkbenchSlotTagsFromRows() {
    var seen = {};
    var out = [];
    savedFiles.forEach(function (rec) {
      if (!rec || !rec.id) return;
      _workbenchRowSlotFilterKeys(__batchWorkbenchRows[String(rec.id)] || {}).forEach(function (k) {
        if (!k || seen[k]) return;
        seen[k] = true;
        out.push(k);
      });
    });
    return out;
  }

  function _collectWorkbenchSlotFilterCatalogValues() {
    var out = [];
    var seen = {};
    WORKBENCH_SLOT_FILTER_GROUPS.forEach(function (grp) {
      (grp.items || []).forEach(function (it) {
        var v = _normalizeWorkbenchSlotTag(it && it.value != null ? it.value : it);
        if (!v || seen[v]) return;
        seen[v] = true;
        out.push(v);
      });
    });
    return out;
  }

  function _workbenchSlotFilterIsAllSelected() {
    if (!batchWorkbenchFilterSlotBox) return false;
    var cbs = batchWorkbenchFilterSlotBox.querySelectorAll('input[data-wb-slot-tag]');
    if (!cbs.length) return false;
    var n = 0;
    for (var i = 0; i < cbs.length; i++) {
      if (cbs[i].checked) n++;
    }
    return n >= cbs.length;
  }

  function _updateWorkbenchSlotFilterToggleLabel() {
    if (!batchWorkbenchFilterSlotToggle) return;
    var n = __wbFilterSlotTags.length;
    if (!n) {
      batchWorkbenchFilterSlotToggle.textContent = '签字位 ▼';
      return;
    }
    if (_workbenchSlotFilterIsAllSelected()) {
      batchWorkbenchFilterSlotToggle.textContent = '签字位：全部 ▼';
      return;
    }
    batchWorkbenchFilterSlotToggle.textContent = '签字位：' + n + ' 项 ▼';
  }

  function _readWorkbenchFilterSlotTagsFromUi() {
    if (!batchWorkbenchFilterSlotBox) return [];
    var out = [];
    var cbs = batchWorkbenchFilterSlotBox.querySelectorAll('input[data-wb-slot-tag]');
    for (var i = 0; i < cbs.length; i++) {
      if (!cbs[i].checked) continue;
      var v = _normalizeWorkbenchSlotTag(cbs[i].getAttribute('data-wb-slot-tag'));
      if (v) out.push(v);
    }
    return out;
  }

  function _setWorkbenchFilterSlotChecks(tagKeys) {
    if (!batchWorkbenchFilterSlotBox) return;
    var keys = tagKeys || {};
    var cbs = batchWorkbenchFilterSlotBox.querySelectorAll('input[data-wb-slot-tag]');
    for (var i = 0; i < cbs.length; i++) {
      var v = _normalizeWorkbenchSlotTag(cbs[i].getAttribute('data-wb-slot-tag'));
      cbs[i].checked = !!keys[v];
    }
    __wbFilterSlotTags = _readWorkbenchFilterSlotTagsFromUi();
    _updateWorkbenchSlotFilterToggleLabel();
  }

  var __wbSlotFilterUiBuilt = false;
  var __wbSlotFilterExtrasSig = '';
  var __wbSlotFilterOnChange = null;

  function _workbenchSlotFilterExtrasSignature(rowTags) {
    return (rowTags || [])
      .slice()
      .sort()
      .join('\u0001');
  }

  function _syncWorkbenchSlotFilterAvailableTags(rowTags) {
    var catalogValues = _collectWorkbenchSlotFilterCatalogValues();
    __wbAvailableSlotTags = catalogValues.slice();
    (rowTags || []).forEach(function (t) {
      t = _normalizeWorkbenchSlotTag(t);
      if (t && __wbAvailableSlotTags.indexOf(t) < 0) __wbAvailableSlotTags.push(t);
    });
    var allowed = {};
    __wbAvailableSlotTags.forEach(function (x) {
      allowed[x] = true;
    });
    __wbFilterSlotTags = __wbFilterSlotTags.filter(function (x) {
      return allowed[x];
    });
  }

  /** 签字位筛选 DOM 只构建一次；勿在 renderBatchWorkbenchTable 里重复调用 */
  function _refreshWorkbenchSlotFilterOptions(opt) {
    if (!batchWorkbenchFilterSlotBox) return;
    opt = opt && typeof opt === 'object' ? opt : {};
    var rowTags = _collectWorkbenchSlotTagsFromRows();
    _syncWorkbenchSlotFilterAvailableTags(rowTags);
    var extrasSig = _workbenchSlotFilterExtrasSignature(rowTags);
    if (__wbSlotFilterUiBuilt && !opt.force && extrasSig === __wbSlotFilterExtrasSig) {
      _updateWorkbenchSlotFilterToggleLabel();
      return;
    }
    __wbSlotFilterExtrasSig = extrasSig;
    var selected = {};
    __wbFilterSlotTags.forEach(function (x) {
      selected[x] = true;
    });
    if (!__wbSlotFilterOnChange) {
      __wbSlotFilterOnChange = function () {
        __wbFilterSlotTags = _readWorkbenchFilterSlotTagsFromUi();
        _updateWorkbenchSlotFilterToggleLabel();
        renderBatchWorkbenchTable();
      };
    }
    var slotGroups = _buildWorkbenchSlotFilterGroupsForUi(rowTags);
    _renderGroupedWorkbenchFilterBox(
      batchWorkbenchFilterSlotBox,
      slotGroups.map(function (g) {
        return {
          title: g.title,
          items: g.items.map(function (it) {
            return {
              value: it.value,
              label: it.label || it.value,
              checked: !!selected[it.value],
            };
          }),
        };
      }),
      'data-wb-slot-tag',
      __wbSlotFilterOnChange
    );
    __wbSlotFilterUiBuilt = true;
    _setWorkbenchFilterSlotChecks(
      __wbFilterSlotTags.reduce(function (m, x) {
        m[x] = true;
        return m;
      }, {})
    );
  }

  function _selectAllWorkbenchSlotFilters() {
    if (!batchWorkbenchFilterSlotBox) return;
    var cbs = batchWorkbenchFilterSlotBox.querySelectorAll('input[data-wb-slot-tag]');
    for (var i = 0; i < cbs.length; i++) cbs[i].checked = true;
    __wbFilterSlotTags = _readWorkbenchFilterSlotTagsFromUi();
    _updateWorkbenchSlotFilterToggleLabel();
    renderBatchWorkbenchTable();
  }

  function _clearWorkbenchSlotFilters() {
    if (!batchWorkbenchFilterSlotBox) return;
    var cbs = batchWorkbenchFilterSlotBox.querySelectorAll('input[data-wb-slot-tag]');
    for (var i = 0; i < cbs.length; i++) cbs[i].checked = false;
    __wbFilterSlotTags = [];
    _updateWorkbenchSlotFilterToggleLabel();
    renderBatchWorkbenchTable();
  }

  function _toggleWorkbenchSlotFilterPanel(open) {
    if (!batchWorkbenchFilterSlotWrap) return;
    if (open) batchWorkbenchFilterSlotWrap.classList.add('open');
    else batchWorkbenchFilterSlotWrap.classList.remove('open');
  }

  function _workbenchRowMatchesSlotFilter(row) {
    if (!__wbFilterSlotTags.length) return false;
    if (_workbenchSlotFilterIsAllSelected()) return true;
    var rowKeys = _workbenchRowSlotFilterKeys(row);
    if (!rowKeys.length) return false;
    for (var i = 0; i < __wbFilterSlotTags.length; i++) {
      if (_slotFilterValueMatchesRow(__wbFilterSlotTags[i], rowKeys)) return true;
    }
    return false;
  }

  function _setWorkbenchFilterDetectIssuesOnly() {
    _setWorkbenchFilterStatusChecks({
      识别有误: true,
      识别失败: true,
      识别超时: true,
      未识别: true,
    });
    renderBatchWorkbenchTable();
  }

  function _setWorkbenchFilterNotSignableOnly() {
    _setWorkbenchFilterStatusChecks({});
    _setWorkbenchFilterSlotChecks({ 不可签: true });
    renderBatchWorkbenchTable();
  }

  function _workbenchRowMatchesStatusFilter(row) {
    if (!__wbFilterStatuses.length) return false;
    if (_workbenchStatusFilterIsAllSelected()) return true;
    var shown = row ? String(row.status || '') : '';
    var bucket = _workbenchStatusBucket(shown);
    for (var i = 0; i < __wbFilterStatuses.length; i++) {
      var f = __wbFilterStatuses[i];
      if (f === '__detect_issues__') {
        if (_isWorkbenchDetectIssueStatus(shown)) return true;
      } else if (f === bucket) return true;
    }
    return false;
  }

  function selectWorkbenchRowsByStatus(selected, opt) {
    opt = opt && typeof opt === 'object' ? opt : {};
    if (!__wbFilterStatuses.length) return 0;
    var on = !!selected;
    var merge = opt.merge !== false;
    var n = 0;
    savedFiles.forEach(function (rec) {
      if (!rec || !rec.id) return;
      var row = __batchWorkbenchRows[String(rec.id)];
      if (!row || !_workbenchRowMatchesStatusFilter(row)) return;
      if (on && merge && row.selected) return;
      row.selected = on;
      n++;
    });
    syncHiddenBatchPicks();
    renderBatchWorkbenchTable();
    return n;
  }

  function _pinWorkbenchFilterNameTerm() {
    var t = String(__wbFilterNameDraft || '').trim();
    if (!t) return false;
    if (__wbFilterNameTerms.indexOf(t) < 0) __wbFilterNameTerms.push(t);
    __wbFilterNameDraft = '';
    if (batchWorkbenchFilterName) batchWorkbenchFilterName.value = '';
    _renderWorkbenchFilterNameTags();
    renderBatchWorkbenchTable();
    return true;
  }

  function _renderWorkbenchFilterNameTags() {
    if (!batchWorkbenchFilterNameTags) return;
    if (!__wbFilterNameTerms.length) {
      batchWorkbenchFilterNameTags.style.display = 'none';
      batchWorkbenchFilterNameTags.innerHTML = '';
      return;
    }
    batchWorkbenchFilterNameTags.style.display = 'flex';
    batchWorkbenchFilterNameTags.innerHTML = '';
    var label = document.createElement('span');
    label.style.color = 'var(--text-muted)';
    label.textContent = '已筛选文件名：';
    batchWorkbenchFilterNameTags.appendChild(label);
    __wbFilterNameTerms.forEach(function (term) {
      var tag = document.createElement('span');
      tag.className = 'wb-filter-tag';
      var txt = document.createElement('span');
      txt.textContent = term;
      tag.appendChild(txt);
      var rm = document.createElement('button');
      rm.type = 'button';
      rm.setAttribute('aria-label', '移除');
      rm.textContent = '×';
      rm.addEventListener('click', function () {
        var ix = __wbFilterNameTerms.indexOf(term);
        if (ix >= 0) __wbFilterNameTerms.splice(ix, 1);
        _renderWorkbenchFilterNameTags();
        renderBatchWorkbenchTable();
      });
      tag.appendChild(rm);
      batchWorkbenchFilterNameTags.appendChild(tag);
    });
  }

  function _clearWorkbenchFilters() {
    __wbFilterNameTerms = [];
    __wbFilterNameDraft = '';
    __wbFilterStatuses = [];
    __wbFilterSlotTags = [];
    if (batchWorkbenchFilterName) batchWorkbenchFilterName.value = '';
    _setWorkbenchFilterStatusChecks({});
    _setWorkbenchFilterSlotChecks({});
    _renderWorkbenchFilterNameTags();
    renderBatchWorkbenchTable();
  }

  function _workbenchHasActiveFilter() {
    return __wbFilterNameTerms.length > 0 || __wbFilterStatuses.length > 0 || __wbFilterSlotTags.length > 0;
  }

  function _workbenchNameFilterTerms() {
    return __wbFilterNameTerms.slice();
  }

  function _countWorkbenchSelectedRows() {
    var n = 0;
    savedFiles.forEach(function (rec) {
      if (!rec || !rec.id) return;
      var row = __batchWorkbenchRows[String(rec.id)];
      if (row && row.selected) n++;
    });
    return n;
  }

  function _setWorkbenchRowsSelected(on, onlyFilteredVisible) {
    savedFiles.forEach(function (r) {
      if (!r || !r.id) return;
      var row = __batchWorkbenchRows[String(r.id)];
      if (!row) return;
      if (onlyFilteredVisible && !_workbenchRowPassesFilter(r, row)) return;
      row.selected = !!on;
    });
  }

  function _canvasOverrideForRole(fileId, rid) {
    var st = fileUiCache[String(fileId)] || {};
    var co = (st.canvasOverrides || {})[rid] || {};
    var sig = co.sig || '';
    var date = co.date || '';
    if (String(selectedFileId) === String(fileId)) {
      var sigC = document.getElementById('sig_' + rid);
      var dateC = document.getElementById('date_' + rid);
      if (sigC && !isCanvasBlank(sigC)) sig = _normalizedPngDataUrl(sigC, 'sig');
      if (dateC && !isCanvasBlank(dateC)) date = _normalizedPngDataUrl(dateC, 'date');
    }
    return { sig: sig, date: date };
  }

  function roleCanvasOverrideReady(fileId, rid) {
    var ov = _canvasOverrideForRole(fileId, rid);
    return !!(ov.sig && ov.date);
  }

  function assessFileMaterialStatus(fileId) {
    var roles = mergeDetectedRolesForFile(fileId);
    if (!roles.length) {
      return {
        state: 'unknown',
        label: '未识别',
        title: '请先识别签字位',
        matched: 0,
        canvas: 0,
        total: 0,
        missingRoleIds: [],
        canvasRoleIds: [],
      };
    }
    var map = (fileUiCache[String(fileId)] || {}).currentRoleMap || {};
    var matched = [];
    var sigOnly = [];
    var dateOnly = [];
    var canvasOnly = [];
    var missing = [];
    var roleDetails = [];
    roles.forEach(function (r) {
      var rid = r.id;
      var p = map[rid];
      var detail = { id: rid, label: roleLabel(rid), state: 'missing' };
      if (roleLibraryMaterialReady(p)) {
        matched.push(rid);
        detail.state = 'full';
      } else if (roleLibrarySigReady(p)) {
        matched.push(rid);
        sigOnly.push(rid);
        detail.state = 'sig_only';
      } else if (roleLibraryDateReady(p)) {
        matched.push(rid);
        dateOnly.push(rid);
        detail.state = 'date_only';
      } else if (roleCanvasOverrideReady(fileId, rid)) {
        canvasOnly.push(rid);
        detail.state = 'canvas';
      } else {
        missing.push(rid);
      }
      roleDetails.push(detail);
    });
    var total = roles.length;
    var libN = matched.length;
    var canvasN = canvasOnly.length;
    var titleParts = roleDetails.map(function (d) {
      var st =
        d.state === 'full'
          ? '已匹配'
          : d.state === 'sig_only'
            ? '仅签名'
            : d.state === 'date_only'
              ? '仅日期'
              : d.state === 'canvas'
                ? '画布'
                : '缺素材';
      return d.label + '：' + st;
    });
    var title = titleParts.join('；');
    if (!missing.length) {
      if (libN === total && !sigOnly.length) {
        return {
          state: 'full',
          label: '全部已匹配',
          title: title,
          matched: libN,
          canvas: 0,
          total: total,
          missingRoleIds: [],
          canvasRoleIds: [],
          sigOnlyRoleIds: [],
          dateOnlyRoleIds: [],
          roleDetails: roleDetails,
        };
      }
      if (canvasN > 0 && libN + canvasN === total) {
        return {
          state: 'canvas',
          label: libN ? libN + '/' + total + ' 库+' + canvasN + ' 画布' : '画布已补全',
          title: title,
          matched: libN,
          canvas: canvasN,
          total: total,
          missingRoleIds: [],
          canvasRoleIds: canvasOnly,
          sigOnlyRoleIds: sigOnly,
          dateOnlyRoleIds: dateOnly,
          roleDetails: roleDetails,
        };
      }
    }
    if (libN === 0 && canvasN === 0) {
      return {
        state: 'none',
        label: '均未就绪',
        title: title,
        matched: 0,
        canvas: 0,
        total: total,
        missingRoleIds: missing,
        canvasRoleIds: [],
        sigOnlyRoleIds: [],
        dateOnlyRoleIds: dateOnly,
        roleDetails: roleDetails,
      };
    }
    return {
      state: 'partial',
      label: libN + '/' + total + ' 已匹配',
      title: title,
      matched: libN,
      canvas: canvasN,
      total: total,
      missingRoleIds: missing,
      canvasRoleIds: canvasOnly,
      sigOnlyRoleIds: sigOnly,
      dateOnlyRoleIds: dateOnly,
      roleDetails: roleDetails,
    };
  }

  function refreshWorkbenchRowMaterial(fileId) {
    var row = __batchWorkbenchRows[String(fileId)];
    if (!row) return;
    row.material = assessFileMaterialStatus(fileId);
  }

  function setBatchWorkbenchAdvanced(on) {
    __batchWorkbenchAdvancedOpen = !!on;
    try {
      if (on) document.body.classList.add('batch-advanced-mode');
      else document.body.classList.remove('batch-advanced-mode');
    } catch (_) {}
    if (batchWorkbenchAdvancedCb) batchWorkbenchAdvancedCb.checked = !!on;
    if (on && !__batchWorkbenchEditFileId && savedFiles.length) {
      var pick = null;
      for (var i = 0; i < savedFiles.length; i++) {
        var r = savedFiles[i];
        if (!r || !r.id) continue;
        var a = assessFileMaterialStatus(r.id);
        if (a.missingRoleIds && a.missingRoleIds.length) {
          pick = r.id;
          break;
        }
      }
      selectWorkbenchFileForAdvanced(pick || savedFiles[0].id);
    }
    if (!on) {
      saveCurrentFileCanvasToCache(selectedFileId);
      renderBatchWorkbenchTable();
    }
  }

  function updateNeedSignCardTitleForBatch(fileId) {
    var h2 = document.getElementById('needSignCardTitle');
    var hint = document.getElementById('needSignCardHint');
    if (!h2) return;
    if (!isBatchWorkbenchMode()) {
      h2.textContent = '本文件需签角色（检测结果）';
      if (hint) {
        hint.textContent =
          '根据模板中的编制/审核/批准等关键词自动识别；为每个角色绑定签名/日期素材后可「载入签名」「载入日期」到下方画布。表格较宽时可左右滑动查看。';
      }
      return;
    }
    var row = __batchWorkbenchRows[String(fileId)];
    h2.textContent = row
      ? '当前文件：' + (row.name || fileId)
      : '本文件需签角色（检测结果）';
    if (hint) {
      hint.textContent =
        '与单文件页相同：选择签署人、签名/日期素材（自动匹配后默认选中）。修改会保存到当前文件。';
    }
  }

  function selectWorkbenchFileForMaterialEdit(fileId, opt) {
    opt = opt && typeof opt === 'object' ? opt : {};
    if (!fileId) return;
    if (
      __batchWorkbenchEditFileId &&
      String(__batchWorkbenchEditFileId) !== String(fileId)
    ) {
      if (__batchWorkbenchAdvancedOpen) {
        saveCurrentFileCanvasToCache(__batchWorkbenchEditFileId);
      }
      saveCurrentFileUiToCache(__batchWorkbenchEditFileId);
    }
    __batchWorkbenchEditFileId = fileId;
    if (String(selectedFileId) !== String(fileId)) {
      saveCurrentFileUiToCache(selectedFileId);
      selectedFileId = fileId;
    }
    syncRowToHandoffCtx(fileId);
    var st = fileUiCache[fileId] || {};
    lastDetectData = st.lastDetectData || null;
    lastDetectFileId = st.lastDetectData ? fileId : null;
    lastDetectError = st.lastDetectError || '';
    currentRoleMap = _deepCloneJsonish(st.currentRoleMap || {});
    if (!restoreFileUiFromCache(fileId)) {
      resetAllRoleChecks();
      if (!isBatchWorkbenchMode()) renderNeedSignTable();
    }
    updateNeedSignCardTitleForBatch(fileId);
    refreshWorkbenchRowMaterial(fileId);
    renderBatchWorkbenchTable();
  }

  function saveCurrentFileCanvasToCache(fileId) {
    if (!fileId || String(selectedFileId) !== String(fileId)) return;
    fileUiCache[fileId] = fileUiCache[fileId] || {};
    var co = fileUiCache[fileId].canvasOverrides || {};
    ROLES.forEach(function (r) {
      var rid = r.id;
      var sigC = document.getElementById('sig_' + rid);
      var dateC = document.getElementById('date_' + rid);
      if (!co[rid]) co[rid] = {};
      if (sigC && !isCanvasBlank(sigC)) co[rid].sig = _normalizedPngDataUrl(sigC, 'sig');
      else delete co[rid].sig;
      if (dateC && !isCanvasBlank(dateC)) co[rid].date = _normalizedPngDataUrl(dateC, 'date');
      else delete co[rid].date;
      if (!co[rid].sig && !co[rid].date) delete co[rid];
    });
    fileUiCache[fileId].canvasOverrides = co;
    refreshWorkbenchRowMaterial(fileId);
    if (isBatchWorkbenchMode()) renderBatchWorkbenchTable();
  }

  function restoreFileCanvasFromCache(fileId) {
    ROLES.forEach(function (r) {
      var rid = r.id;
      var sigC = document.getElementById('sig_' + rid);
      var dateC = document.getElementById('date_' + rid);
      if (canvases['sig_' + rid] && canvases['sig_' + rid].clear) canvases['sig_' + rid].clear();
      if (canvases['date_' + rid] && canvases['date_' + rid].clear) canvases['date_' + rid].clear();
    });
    var st = fileUiCache[String(fileId)] || {};
    var co = st.canvasOverrides || {};
    Object.keys(co).forEach(function (rid) {
      var p = co[rid];
      if (p && p.sig) drawUrlToCanvas('sig_' + rid, p.sig, false);
      if (p && p.date) drawUrlToCanvas('date_' + rid, p.date, false);
    });
    resizeCanvasesForRoles(selectedRoleIds());
  }

  function collectWorkbenchCanvasOverridesForFile(fileId) {
    if (String(selectedFileId) === String(fileId)) {
      saveCurrentFileCanvasToCache(fileId);
    }
    var sig_map = {};
    var date_map = {};
    var map = (fileUiCache[String(fileId)] || {}).currentRoleMap || {};
    mergeDetectedRolesForFile(fileId).forEach(function (r) {
      var rid = r.id;
      if (roleLibraryMaterialReady(map[rid])) return;
      var ov = _canvasOverrideForRole(fileId, rid);
      if (ov.sig) sig_map[rid] = ov.sig;
      if (ov.date) date_map[rid] = ov.date;
    });
    return { sig_map: sig_map, date_map: date_map };
  }

  function selectWorkbenchFileForAdvanced(fileId) {
    if (!fileId) return;
    if (!__batchWorkbenchAdvancedOpen) setBatchWorkbenchAdvanced(true);
    selectWorkbenchFileForMaterialEdit(fileId, { scroll: false });
    restoreFileCanvasFromCache(fileId);
    var assess = assessFileMaterialStatus(fileId);
    (assess.missingRoleIds || []).concat(assess.canvasRoleIds || []).forEach(function (rid) {
      setRoleChecked(rid, true);
    });
    requestAnimationFrame(function () {
      resizeCanvasesForRoles(selectedRoleIds());
    });
    renderBatchWorkbenchTable();
    try {
      var card = document.getElementById('batchRoleCanvasCard');
      if (card) card.scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (_) {}
  }

  function setBatchWorkbenchMsg(msg, level) {
    if (!batchWorkbenchMsg) return;
    batchWorkbenchMsg.style.display = msg ? 'block' : 'none';
    var cls = 'btn-inline-feedback';
    if (level === true || level === 'error' || level === 'err') cls += ' is-error';
    else if (level === 'warn' || level === 'warning') cls += ' is-warn';
    else if (level === 'ok' || level === 'success') cls += ' is-ok';
    batchWorkbenchMsg.className = cls;
    batchWorkbenchMsg.textContent = msg || '';
    // 兜底内联样式（防止旧 CSS 缺类时仍能看到颜色）
    batchWorkbenchMsg.style.padding = '8px 12px';
    batchWorkbenchMsg.style.borderRadius = '6px';
    batchWorkbenchMsg.style.marginBottom = '8px';
    batchWorkbenchMsg.style.whiteSpace = 'pre-wrap';
    batchWorkbenchMsg.style.lineHeight = '1.5';
    if (cls.indexOf('is-error') >= 0) {
      batchWorkbenchMsg.style.background = '#ffebee';
      batchWorkbenchMsg.style.color = '#b71c1c';
      batchWorkbenchMsg.style.border = '1px solid #ef9a9a';
    } else if (cls.indexOf('is-warn') >= 0) {
      batchWorkbenchMsg.style.background = '#fff3e0';
      batchWorkbenchMsg.style.color = '#bf360c';
      batchWorkbenchMsg.style.border = '1px solid #ffcc80';
    } else if (cls.indexOf('is-ok') >= 0) {
      batchWorkbenchMsg.style.background = '#e8f5e9';
      batchWorkbenchMsg.style.color = '#1b5e20';
      batchWorkbenchMsg.style.border = '1px solid #a5d6a7';
    } else {
      batchWorkbenchMsg.style.background = '#e3f2fd';
      batchWorkbenchMsg.style.color = '#0d47a1';
      batchWorkbenchMsg.style.border = '1px solid #90caf9';
    }
  }

  function setBatchWorkbenchProgress(done, total, text, keepVisible) {
    var t = Math.max(0, parseInt(total, 10) || 0);
    var d = Math.max(0, parseInt(done, 10) || 0);
    if (t > 0 && d > t) d = t;
    var pct = t > 0 ? Math.round((d / t) * 100) : 0;
    var line = text || (t > 0 ? '进度 ' + d + '/' + t + '（' + pct + '%）' : '处理中…');
    if (isBatchWorkbenchMode()) {
      if (batchWorkbenchProgressWrap && batchWorkbenchProgressBar && batchWorkbenchProgressText) {
        var show = !!keepVisible || t > 0;
        batchWorkbenchProgressWrap.style.display = show ? 'block' : 'none';
        if (show) {
          batchWorkbenchProgressBar.max = t > 0 ? t : 100;
          batchWorkbenchProgressBar.value = t > 0 ? d : 0;
          batchWorkbenchProgressText.textContent = line;
        } else {
          batchWorkbenchProgressText.textContent = '';
        }
      }
      return;
    }
    if (__pageProgressDepth > 0 || keepVisible || t > 0) {
      updatePageProgress(line, { done: d, total: t });
    }
  }

  function workbenchStateFromRow(fileId) {
    var row = __batchWorkbenchRows[String(fileId)];
    if (!row) return null;
    return {
      status: row.status || '',
      rolesLabel: row.rolesLabel || '',
      slotLabel: row.slotLabel || '',
      slotRolesLine: row.slotRolesLine || '',
      slotLayoutLine: row.slotLayoutLine || '',
      slotBadgeClass: row.slotBadgeClass || '',
      slotExplain: row.slotExplain || '',
      slotTags: Array.isArray(row.slotTags) ? row.slotTags.slice() : [],
      detectExplain: row.detectExplain || '',
      detectWrongNote: row.detectWrongNote || '',
      manualDetectWrong: !!row.manualDetectWrong,
      editor: row.editor || '',
      reviewer: row.reviewer || '',
      approver: row.approver || '',
      doc_date: row.doc_date || '',
      locale: row.locale || '',
      country: row.country || '',
      selected: !!row.selected,
    };
  }

  function applyWorkbenchStateFromCache(fileId, wb) {
    if (!fileId || !wb || typeof wb !== 'object') return;
    var row = __batchWorkbenchRows[String(fileId)];
    if (!row) return;
    if (wb.status != null) {
      var stCached = String(wb.status);
      // 刷新页面时勿恢复「识别中/匹配素材」等中间态，否则签字位列会一直显示处理中
      if (!_isWorkbenchRowPipelineBusy(stCached)) {
        row.status = stCached;
      }
    }
    if (typeof wb.manualDetectWrong === 'boolean') {
      row.manualDetectWrong = wb.manualDetectWrong;
    }
    if (wb.detectWrongNote != null) row.detectWrongNote = String(wb.detectWrongNote);
    if (wb.rolesLabel != null) row.rolesLabel = String(wb.rolesLabel);
    if (wb.slotLabel != null) row.slotLabel = String(wb.slotLabel);
    if (wb.slotRolesLine != null) row.slotRolesLine = String(wb.slotRolesLine);
    if (wb.slotLayoutLine != null) row.slotLayoutLine = String(wb.slotLayoutLine);
    if (wb.slotBadgeClass != null) row.slotBadgeClass = String(wb.slotBadgeClass);
    if (wb.slotExplain != null) row.slotExplain = String(wb.slotExplain);
    if (Array.isArray(wb.slotTags)) row.slotTags = wb.slotTags.slice();
    if (wb.detectExplain != null) row.detectExplain = String(wb.detectExplain);
    if (wb.editor != null) row.editor = String(wb.editor);
    if (wb.reviewer != null) row.reviewer = String(wb.reviewer);
    if (wb.approver != null) row.approver = String(wb.approver);
    if (wb.doc_date != null) row.doc_date = String(wb.doc_date);
    if (wb.locale != null) row.locale = String(wb.locale);
    if (wb.country != null) row.country = String(wb.country);
    if (typeof wb.selected === 'boolean') row.selected = wb.selected;
    syncRowToHandoffCtx(fileId);
  }

  var __persistFileCacheTimers = {};
  function _fileSessionCachePayload(fileId, lightOnly) {
    var payload = {};
    if (!lightOnly) {
      var st = fileUiCache[String(fileId)] || {};
      if (st.lastDetectData) payload.detect = st.lastDetectData;
    }
    var wb = workbenchStateFromRow(fileId);
    if (wb) payload.workbench = wb;
    var corr = __fileDetectCorrectionCache[String(fileId)];
    if (corr && typeof corr === 'object' && Object.keys(corr).length) {
      payload.detect_correction = corr;
    }
    return payload;
  }

  function persistFileSessionCacheNow(fileId, lightOnly) {
    if (!fileId) return Promise.resolve();
    var payload = _fileSessionCachePayload(fileId, !!lightOnly);
    if (!Object.keys(payload).length) return Promise.resolve();
    return fetchJson(apiUrl('/api/sign/files/' + encodeURIComponent(fileId) + '/file-cache'), {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
      timeoutMs: lightOnly ? 45000 : 60000,
    }).catch(function () {});
  }

  function _detectCorrectionBatchTimeoutMs(count) {
    var n = Math.max(1, count || 1);
    return Math.min(600000, Math.max(120000, 90000 + n * 2000));
  }

  function flushLightFileSessionCaches(fileIds) {
    var ids = (fileIds || []).map(String).filter(Boolean);
    if (!ids.length) return Promise.resolve();
    var i = 0;
    var chunk = 6;
    function nextChunk() {
      var slice = ids.slice(i, i + chunk);
      if (!slice.length) return Promise.resolve();
      i += chunk;
      return Promise.all(
        slice.map(function (fid) {
          return persistFileSessionCacheNow(fid, true);
        })
      ).then(nextChunk);
    }
    return nextChunk();
  }

  function schedulePersistFileSessionCache(fileId) {
    if (!fileId) return;
    var fid = String(fileId);
    if (__persistFileCacheTimers[fid]) clearTimeout(__persistFileCacheTimers[fid]);
    __persistFileCacheTimers[fid] = setTimeout(function () {
      delete __persistFileCacheTimers[fid];
      persistFileSessionCacheNow(fid);
    }, 400);
  }

  function flushPendingFileSessionCaches(fileIdsOpt) {
    var ids = fileIdsOpt && fileIdsOpt.length
      ? fileIdsOpt.map(String)
      : Object.keys(__persistFileCacheTimers);
    ids.forEach(function (fid) {
      if (__persistFileCacheTimers[fid]) {
        clearTimeout(__persistFileCacheTimers[fid]);
        delete __persistFileCacheTimers[fid];
      }
    });
    var uniq = {};
    ids.forEach(function (fid) {
      if (fid) uniq[fid] = true;
    });
    return Promise.all(
      Object.keys(uniq).map(function (fid) {
        return persistFileSessionCacheNow(fid);
      })
    );
  }

  function _applyOneFileSessionCacheEntry(fid, ent) {
    ent = ent || {};
    if (ent.detect) {
      cacheDetectResultForFile(fid, ent.detect, ent.detect.error || '');
    }
    if (ent.map) {
      cachePatchCurrentRoleMap(fid, ent.map);
      _lastPersistedRoleMapStable[String(fid)] = _roleMapStableJson(ent.map);
    }
    if (ent.detect_correction && typeof ent.detect_correction === 'object') {
      __fileDetectCorrectionCache[fid] = ent.detect_correction;
    }
    if (ent.workbench) {
      applyWorkbenchStateFromCache(fid, ent.workbench);
    } else if (ent.detect) {
      syncWorkbenchRowFromDetect(fid, ent.detect);
    }
    var row = __batchWorkbenchRows[String(fid)];
    if (row) {
      finalizeWorkbenchRowStatus(fid);
    }
  }

  function hydrateFileCachesFromServer(opt) {
    opt = opt && typeof opt === 'object' ? opt : {};
    if (!IS_FILE_SIGN_PAGE) return Promise.resolve();
    if (__fileCacheHydratePromise) return __fileCacheHydratePromise;
    __fileCacheHydratePromise = fetchJsonWithRetry(
      apiUrl('/api/sign/files/file-caches') + '?_=' + Date.now(),
      { timeoutMs: 90000, cache: 'no-store' },
      { maxTry: 2, delayMs: 1500 }
    )
      .then(function (r) {
        var j = (r && r.data) || {};
        if (!j.ok || !j.caches || typeof j.caches !== 'object') return;
        var keys = Object.keys(j.caches);
        if (!keys.length) return;
        var caches = j.caches;
        var idx = 0;
        var chunk = opt.chunkSize > 0 ? opt.chunkSize : 12;
        return new Promise(function (resolve) {
          function step() {
            var end = Math.min(idx + chunk, keys.length);
            for (; idx < end; idx++) {
              _applyOneFileSessionCacheEntry(keys[idx], caches[keys[idx]]);
            }
            if (opt.onProgress) {
              try {
                opt.onProgress(idx, keys.length);
              } catch (_) {}
            }
            if (idx < keys.length) {
              requestAnimationFrame(step);
            } else {
              try {
                _refreshWorkbenchSlotFilterOptions();
                renderBatchWorkbenchTable();
              } catch (_) {}
              resolve();
            }
          }
          requestAnimationFrame(step);
        });
      })
      .catch(function (e) {
        var msg = (e && e.message) || String(e);
        try {
          setBatchWorkbenchMsg(
            '识别/匹配缓存恢复失败（页面仍可使用）：' + msg + '。可点「刷新列表」重试。',
            true
          );
        } catch (_) {}
      })
      .finally(function () {
        __fileCacheHydratePromise = null;
      });
    return __fileCacheHydratePromise;
  }

  function _workbenchFilteredFileIds() {
    var ids = [];
    savedFiles.forEach(function (rec) {
      if (!rec || !rec.id) return;
      var row = __batchWorkbenchRows[String(rec.id)];
      if (!_workbenchRowPassesFilter(rec, row)) return;
      ids.push(String(rec.id));
    });
    return ids;
  }

  function selectWorkbenchFilteredRows(selected, opt) {
    opt = opt && typeof opt === 'object' ? opt : {};
    var on = !!selected;
    var merge = opt.merge !== false;
    var n = 0;
    savedFiles.forEach(function (rec) {
      if (!rec || !rec.id) return;
      var row = __batchWorkbenchRows[String(rec.id)];
      if (!_workbenchRowPassesFilter(rec, row)) return;
      if (!row) return;
      if (on && merge && row.selected) return;
      row.selected = on;
      n++;
    });
    syncHiddenBatchPicks();
    renderBatchWorkbenchTable();
    return n;
  }

  function _workbenchRowPassesFilter(rec, row) {
    if (!rec || !rec.id) return false;
    var nm = String(rec.name || rec.id || '').toLowerCase();
    var nameTerms = _workbenchNameFilterTerms();
    if (nameTerms.length) {
      var nameHit = false;
      for (var ni = 0; ni < nameTerms.length; ni++) {
        var term = String(nameTerms[ni] || '').trim().toLowerCase();
        if (term && nm.indexOf(term) >= 0) {
          nameHit = true;
          break;
        }
      }
      if (!nameHit) return false;
    }
    if (__wbFilterStatuses.length && !_workbenchRowMatchesStatusFilter(row)) return false;
    if (__wbFilterSlotTags.length && !_workbenchRowMatchesSlotFilter(row)) return false;
    return true;
  }

  function _updateWorkbenchFilterHint(shown, total) {
    if (!batchWorkbenchFilterHint) return;
    var sel = _countWorkbenchSelectedRows();
    if (!_workbenchHasActiveFilter()) {
      batchWorkbenchFilterHint.textContent = total
        ? '共 ' + total + ' 个文件' + (sel ? ' · 已勾选 ' + sel + ' 个' : '')
        : '';
      return;
    }
    var parts = ['显示 ' + shown + ' / ' + total];
    if (__wbFilterNameTerms.length) {
      parts.push('文件名[' + __wbFilterNameTerms.join('|') + ']');
    }
    if (__wbFilterStatuses.length) {
      if (_workbenchStatusFilterIsAllSelected()) {
        parts.push('状态[全部]');
      } else {
        parts.push('状态[' + __wbFilterStatuses.join('|') + ']');
      }
    }
    if (__wbFilterSlotTags.length) {
      if (_workbenchSlotFilterIsAllSelected()) {
        parts.push('签字位[全部]');
      } else {
        parts.push('签字位[' + __wbFilterSlotTags.join('|') + ']');
      }
    }
    if (sel) parts.push('已勾选 ' + sel + ' 个');
    batchWorkbenchFilterHint.textContent = parts.join(' · ');
  }

  function _summarizeBatchWorkbenchProgress(scopeFileIds) {
    var ids = (scopeFileIds && scopeFileIds.length
      ? scopeFileIds
      : Object.keys(__batchWorkbenchRows || {})
    ).map(String);
    var s = { total: ids.length, done: 0, ready: 0, partial: 0, waitMatch: 0, detectWrong: 0,
      noSign: 0, detectTimeout: 0, detectFail: 0, noField: 0, busy: 0, pending: 0,
      failItems: [] };
    ids.forEach(function (fid) {
      var r = __batchWorkbenchRows[fid];
      if (!r) return;
      var t = String(r.status || '');
      if (t === '无需签字') { s.noSign++; s.done++; }
      else if (t === '就绪') { s.ready++; s.done++; }
      else if (t === '部分就绪') { s.partial++; s.done++; }
      else if (t === '待匹配') { s.waitMatch++; s.done++; }
      else if (t === '识别有误') { s.detectWrong++; s.done++;
        s.failItems.push({ fid: fid, name: r.name || fid, status: t, detail: r.detectWrongNote || r.rolesLabel || '' }); }
      else if (/识别超时/.test(t)) { s.detectTimeout++; s.done++;
        s.failItems.push({ fid: fid, name: r.name || fid, status: t, detail: r.rolesLabel || '' }); }
      else if (/识别失败/.test(t)) { s.detectFail++; s.done++;
        s.failItems.push({ fid: fid, name: r.name || fid, status: t, detail: r.rolesLabel || '' }); }
      else if (/未识别/.test(t)) { s.noField++; s.done++;
        s.failItems.push({ fid: fid, name: r.name || fid, status: t, detail: r.rolesLabel || '' }); }
      else if (/识别中|匹配素材|识别…|分析中/.test(t)) { s.busy++; }
      else { s.pending++; }
    });
    return s;
  }

  function _csvCell(v) {
    var s = String(v == null ? '' : v);
    if (/[",\r\n]/.test(s)) {
      return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  }

  function _workbenchIssueSuggestion(status, mat, details) {
    var s = String(status || '');
    if (/识别超时/.test(s)) {
      return '单文件点「识别」重试；必要时提高 SIGN_DETECT_TIMEOUT_MS';
    }
    if (/识别失败/.test(s)) {
      return '单文件点「识别」重试，并检查文档是否可正常打开';
    }
    if (/无需签字/.test(s)) {
      return '本类文档按规则无需签字，可取消勾选后批量处理其余文件';
    }
    if (/未识别/.test(s)) {
      return '检查模板签批栏标签（编写/审核/批准）是否规范，或修正 JSON 规则后批量重新识别';
    }
    if (/识别有误/.test(s)) {
      return '人工标记识别有误：保存纠正登记后点「批量识别（角色+签字位）」重新识别';
    }
    if (/待匹配/.test(s)) {
      return '识别结果可用，点「批量匹配素材」或补全库素材';
    }
    if (mat && mat.missingRoleIds && mat.missingRoleIds.length) {
      return '为缺素材角色补签名/日期素材，或开启高级模式画布补签';
    }
    if ((mat && mat.sigOnlyRoleIds && mat.sigOnlyRoleIds.length) || /仅签名/.test(details || '')) {
      return '补齐日期素材或填写文档日期后重匹配';
    }
    if ((mat && mat.dateOnlyRoleIds && mat.dateOnlyRoleIds.length) || /仅日期/.test(details || '')) {
      return '补齐签名素材后重匹配';
    }
    return '已就绪可直接批量签字';
  }

  function _collectBatchWorkbenchIssueRows(onlySelected) {
    var rows = [];
    savedFiles.forEach(function (rec) {
      if (!rec || !rec.id) return;
      var fid = String(rec.id);
      var row = __batchWorkbenchRows[fid];
      if (!row) return;
      if (onlySelected && !row.selected) return;
      refreshWorkbenchRowMaterial(fid);
      var mat = row.material || assessFileMaterialStatus(fid);
      var status = String(row.status || '');
      var roleText = String(row.rolesLabel || '—');
      var detailParts = [];
      var missingLabels = (mat.missingRoleIds || []).map(roleLabel);
      var sigOnlyLabels = (mat.sigOnlyRoleIds || []).map(roleLabel);
      var dateOnlyLabels = (mat.dateOnlyRoleIds || []).map(roleLabel);
      if (missingLabels.length) detailParts.push('缺素材：' + missingLabels.join('、'));
      if (sigOnlyLabels.length) detailParts.push('仅签名：' + sigOnlyLabels.join('、'));
      if (dateOnlyLabels.length) detailParts.push('仅日期：' + dateOnlyLabels.join('、'));
      var details = detailParts.join('；');
      var hasIssue =
        /识别超时|识别失败|未识别|识别有误|部分就绪/.test(status) ||
        !!missingLabels.length ||
        !!sigOnlyLabels.length ||
        !!dateOnlyLabels.length;
      if (!hasIssue) return;
      rows.push({
        file_id: fid,
        file_name: String(row.name || rec.name || fid),
        status: status || '—',
        roles: roleText,
        details: details || roleText || '—',
        suggestion: _workbenchIssueSuggestion(status, mat, details),
      });
    });
    return rows;
  }

  function exportBatchWorkbenchIssues() {
    var hasSelected = savedFiles.some(function (rec) {
      var row = rec && rec.id ? __batchWorkbenchRows[String(rec.id)] : null;
      return !!(row && row.selected);
    });
    var rows = _collectBatchWorkbenchIssueRows(hasSelected);
    if (!rows.length) {
      setBatchWorkbenchMsg('当前未发现异常项，所有文件可直接批量签字。', 'ok');
      return;
    }
    var headers = ['file_id', 'file_name', 'status', 'roles', 'details', 'suggestion'];
    var lines = [headers.join(',')];
    rows.forEach(function (r) {
      lines.push(
        [
          _csvCell(r.file_id),
          _csvCell(r.file_name),
          _csvCell(r.status),
          _csvCell(r.roles),
          _csvCell(r.details),
          _csvCell(r.suggestion),
        ].join(',')
      );
    });
    var ts = new Date();
    var stamp = String(ts.getFullYear()) +
      String(ts.getMonth() + 1).padStart(2, '0') +
      String(ts.getDate()).padStart(2, '0') + '_' +
      String(ts.getHours()).padStart(2, '0') +
      String(ts.getMinutes()).padStart(2, '0') +
      String(ts.getSeconds()).padStart(2, '0');
    var blob = new Blob(['\ufeff' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'batch_sign_issues_' + stamp + '.csv';
    a.click();
    setTimeout(function () {
      try {
        URL.revokeObjectURL(a.href);
      } catch (_) {}
    }, 1500);
    setBatchWorkbenchMsg(
      '已导出问题清单（' + rows.length + ' 条）。建议先按清单补素材，再点击「批量匹配素材」与「一键批量签字」。',
      'ok'
    );
  }

  function enableBatchWorkbenchMode() {
    try {
      document.body.classList.add('batch-workbench-mode');
    } catch (_) {}
    if (batchWorkbenchCard) batchWorkbenchCard.style.display = 'block';
    if (signSourceMode) signSourceMode.value = 'library';
    if (batchModeCb) {
      batchModeCb.checked = true;
      batchModeCb.disabled = !signersDbShare;
    }
    syncLibraryRolesModeRow();
  }

  /** 工作台打开或刷新后，合并服务端文件列表到表格行（保留已有行状态） */
  function mergeBatchWorkbenchRowsFromSavedFiles() {
    savedFiles.forEach(function (rec) {
      if (!rec || !rec.id) return;
      var fid = String(rec.id);
      var existing = __batchWorkbenchRows[fid];
      if (existing) {
        existing.name = String(rec.name || fid);
        return;
      }
      var isHandoff = _isAiwordHandoffFileId(fid);
      var ctx = isHandoff ? __aiwordHandoffCtxByFileId[fid] || {} : {};
      var ed = String(ctx.editor || ctx.writer || '').trim();
      var useEn = _isLikelyEnglishCountry(ctx.country);
      __batchWorkbenchRows[fid] = {
        fileId: fid,
        name: String(rec.name || fid),
        selected: false,
        editor: ed,
        reviewer: String(ctx.reviewer || '').trim(),
        approver: String(ctx.approver || '').trim(),
        doc_date: _formatDateInputValue(ctx.doc_date),
        locale: useEn ? 'en' : 'zh',
        country: String(ctx.country || '').trim(),
        status: isHandoff ? '待处理' : '—',
        rolesLabel: '—',
        slotLabel: '—',
        slotTags: [],
      };
    });
  }

  function refreshWorkbenchFilesFromServer(opt) {
    opt = opt || {};
    if (!IS_FILE_SIGN_PAGE) return Promise.resolve();
    return refreshFileList({
      force: true,
      silent: !!opt.silent,
      softFail: true,
      progressLabel: opt.progressLabel || '正在刷新文件列表…',
      skipPageProgress: !!opt.skipPageProgress,
    })
      .then(function () {
        mergeBatchWorkbenchRowsFromSavedFiles();
        if (savedFiles.length) enableBatchWorkbenchMode();
        return hydrateFileCachesFromServer();
      })
      .then(function () {
        if (isBatchWorkbenchMode()) {
          syncHiddenBatchPicks();
          _refreshWorkbenchSlotFilterOptions();
          renderBatchWorkbenchTable();
        }
      });
  }

  function _formatDateInputValue(iso) {
    return _parseAiwordDocDateIso(iso) || '';
  }

  function rowCtxFromRowState(row) {
    if (!row) return {};
    return {
      editor: String(row.editor || '').trim(),
      writer: String(row.editor || '').trim(),
      reviewer: String(row.reviewer || '').trim(),
      approver: String(row.approver || '').trim(),
      doc_date: _parseAiwordDocDateIso(row.doc_date || ''),
      country: String(row.country || '').trim(),
    };
  }

  function syncRowToHandoffCtx(fileId) {
    var row = __batchWorkbenchRows[String(fileId)];
    if (!row) return;
    var ctx = rowCtxFromRowState(row);
    __aiwordHandoffCtxByFileId[String(fileId)] = ctx;
    if (row.locale === 'en') {
      ctx.country = ctx.country || 'United States';
    } else if (row.locale === 'zh') {
      ctx.country = ctx.country || '中国';
    }
  }

  function initBatchRowsFromHandoff() {
    __batchWorkbenchRows = {};
    savedFiles.forEach(function (rec) {
      if (!rec || !rec.id) return;
      var fid = String(rec.id);
      var isHandoff = _isAiwordHandoffFileId(fid);
      var ctx = isHandoff ? __aiwordHandoffCtxByFileId[fid] || {} : {};
      var ed = String(ctx.editor || ctx.writer || '').trim();
      var useEn = _isLikelyEnglishCountry(ctx.country);
      __batchWorkbenchRows[fid] = {
        fileId: fid,
        name: String(rec.name || fid),
        selected: isHandoff,
        editor: ed,
        reviewer: String(ctx.reviewer || '').trim(),
        approver: String(ctx.approver || '').trim(),
        doc_date: _formatDateInputValue(ctx.doc_date),
        locale: useEn ? 'en' : 'zh',
        country: String(ctx.country || '').trim(),
        status: isHandoff ? '待处理' : '—',
        rolesLabel: '—',
        slotLabel: '—',
        slotTags: [],
      };
    });
  }

  function formatDetectedRolesLabelForFile(fileId) {
    var roles = mergeDetectedRolesForFile(fileId);
    if (!roles.length) return '—';
    return roles
      .map(function (r) {
        return r && r.id ? roleLabel(r.id) : '';
      })
      .filter(Boolean)
      .join('、');
  }

  function syncHiddenBatchPicks() {
    if (!batchWorkbenchHiddenPicks) return;
    batchWorkbenchHiddenPicks.innerHTML = '';
    savedFiles.forEach(function (rec) {
      if (!rec || !rec.id) return;
      var fid = String(rec.id);
      var row = __batchWorkbenchRows[fid];
      var cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.className = 'batch-pick';
      cb.setAttribute('data-id', fid);
      cb.checked = row ? !!row.selected : _isAiwordHandoffFileId(fid);
      cb.addEventListener('change', updateBatchUi);
      batchWorkbenchHiddenPicks.appendChild(cb);
    });
    if (fileListEl && !isBatchWorkbenchMode()) {
      fileListEl.innerHTML = '';
      savedFiles.forEach(function (rec) {
        if (!rec || !rec.id) return;
        var cb2 = document.createElement('input');
        cb2.type = 'checkbox';
        cb2.className = 'batch-pick';
        cb2.setAttribute('data-id', rec.id);
        var row2 = __batchWorkbenchRows[String(rec.id)];
        cb2.checked = row2 ? !!row2.selected : _isAiwordHandoffFileId(rec.id);
        cb2.style.display = 'none';
        fileListEl.appendChild(cb2);
      });
    }
    updateBatchUi();
  }

  function renderBatchWorkbenchTable() {
    if (!batchWorkbenchBody) return;
    batchWorkbenchBody.innerHTML = '';
    var total = 0;
    var shown = 0;
    savedFiles.forEach(function (rec) {
      if (!rec || !rec.id) return;
      total++;
      var fid = String(rec.id);
      if (!__batchWorkbenchRows[fid]) {
        var isHandoff = _isAiwordHandoffFileId(fid);
        var ctx = isHandoff ? __aiwordHandoffCtxByFileId[fid] || {} : {};
        var ed = String(ctx.editor || ctx.writer || '').trim();
        __batchWorkbenchRows[fid] = {
          fileId: fid,
          name: String(rec.name || fid),
          selected: isHandoff,
          editor: ed,
          reviewer: String(ctx.reviewer || '').trim(),
          approver: String(ctx.approver || '').trim(),
          doc_date: _formatDateInputValue(ctx.doc_date),
          locale: _isLikelyEnglishCountry(ctx.country) ? 'en' : 'zh',
          country: String(ctx.country || '').trim(),
          status: isHandoff ? '待处理' : '—',
          rolesLabel: '—',
          slotLabel: '—',
          slotTags: [],
        };
      }
      var row = __batchWorkbenchRows[fid];
      if (!row) return;
      if (!_workbenchRowPassesFilter(rec, row)) return;
      shown++;
      refreshWorkbenchRowMaterial(fid);
      syncWorkbenchRowFromDetect(fid);
      var mat = row.material || assessFileMaterialStatus(fid);
      var tr = document.createElement('tr');
      tr.className = 'batch-wb-row' + (row.selected ? ' row-selected' : '');
      if (String(__batchWorkbenchEditFileId) === fid) tr.classList.add('row-editing');
      var _stTxt = String(row.status || '');
      if (_stTxt === '就绪') tr.classList.add('row-ready');
      if (/识别中|匹配素材|识别…|分析中/.test(_stTxt)) tr.classList.add('row-busy');
      if (/失败|超时|未识别|识别有误/.test(_stTxt)) tr.classList.add('row-error');

      var td0 = document.createElement('td');
      td0.className = 'col-check';
      var chk = document.createElement('input');
      chk.type = 'checkbox';
      chk.checked = !!row.selected;
      chk.addEventListener('change', function () {
        row.selected = chk.checked;
        syncHiddenBatchPicks();
        renderBatchWorkbenchTable();
      });
      td0.appendChild(chk);

      var tdName = document.createElement('td');
      tdName.className = 'col-name';
      tdName.textContent = row.name || fid;
      tdName.title = row.name || fid;

      ensureFileRoleLocales(fid, row);
      var tdDocDate = mkWorkbenchDocDateCell(fid, row);
      var tdLoc = document.createElement('td');
      var locSel = document.createElement('select');
      locSel.innerHTML = '<option value="zh">中文</option><option value="en">英文</option>';
      locSel.value = row.locale === 'en' ? 'en' : 'zh';
      locSel.addEventListener('click', function (ev) {
        ev.stopPropagation();
      });
      locSel.addEventListener('change', function (ev) {
        ev.stopPropagation();
        row.locale = locSel.value;
        syncRowToHandoffCtx(fid);
        var rl = row.locale === 'en' ? 'en' : row.locale === 'zh' ? 'zh' : 'auto';
        var st = fileUiCache[fid] || {};
        st.roleLocales = { author: rl, reviewer: rl, approver: rl };
        fileUiCache[fid] = st;
        ['author', 'reviewer', 'approver'].forEach(function (rid) {
          roleLocaleMap[rid] = rl;
        });
        workbenchRematchFileMaterials(fid);
      });
      tdLoc.className = 'col-locale';
      tdLoc.appendChild(locSel);
      var tdEdSig = mkWorkbenchRoleSigCell(fid, 'author', row);
      var tdEdDate = mkWorkbenchRoleDateCell(fid, 'author', row);
      var tdRvSig = mkWorkbenchRoleSigCell(fid, 'reviewer', row);
      var tdRvDate = mkWorkbenchRoleDateCell(fid, 'reviewer', row);
      var tdApSig = mkWorkbenchRoleSigCell(fid, 'approver', row);
      var tdApDate = mkWorkbenchRoleDateCell(fid, 'approver', row);

      var tdMat = document.createElement('td');
      tdMat.className = 'col-material';
      var badge = document.createElement('span');
      badge.className = 'batch-wb-mat-badge mat-' + (mat.state || 'unknown');
      badge.textContent = mat.label || '—';
      if (mat.title) badge.title = mat.title;
      tdMat.appendChild(badge);

      var tdRoles = document.createElement('td');
      tdRoles.className = 'col-roles';
      var _rolesTxt = row.rolesLabel || '—';
      var _slotExplain = row.slotExplain || '';
      var _detExplain = row.detectExplain || '';
      tdRoles.textContent = _rolesTxt;
      tdRoles.title = _detExplain ? (_rolesTxt + '\n' + _detExplain) : _rolesTxt;

      var tdSlot = document.createElement('td');
      _renderWorkbenchSlotCell(tdSlot, row);

      var tdSt = document.createElement('td');
      tdSt.className = 'col-status';
      var _stShown = row.status || '—';
      tdSt.textContent = _stShown;
      // 完整错误/原因显示在 tooltip，避免列宽吞掉用户能定位问题的关键信息
      tdSt.title = _stShown;
      if (_rolesTxt && _rolesTxt !== '—') tdSt.title += '\n需签：' + _rolesTxt;
      if (_slotExplain) tdSt.title += '\n' + _slotExplain;
      if (_detExplain) tdSt.title += '\n' + _detExplain;
      if (/失败|超时|未识别|识别有误/.test(_stShown)) {
        tdSt.style.color = '#c62828';
        tdSt.style.fontWeight = '600';
      } else if (/识别中|匹配素材|识别…|分析中/.test(_stShown)) {
        tdSt.style.color = '#1a73e8';
      } else if (_stShown === '就绪') {
        tdSt.style.color = '#1b5e20';
        tdSt.style.fontWeight = '600';
      } else if (_stShown === '待匹配') {
        tdSt.style.color = '#e65100';
      }

      var tdOp = document.createElement('td');
      tdOp.className = 'col-op';
      var btnDet = document.createElement('button');
      btnDet.type = 'button';
      btnDet.className = 'btn btn-secondary';
      btnDet.style.padding = '4px 8px';
      btnDet.style.fontSize = '0.82rem';
      btnDet.textContent = '识别';
      btnDet.title = '重新识别需签角色与签字位版式（不匹配素材库）';
      btnDet.addEventListener('click', function () {
        // 强制重新识别：清掉缓存与"已处理"标记，确保真的重跑而不是直接命中缓存
        var st = fileUiCache[fid] || {};
        st.lastDetectData = null;
        st.lastDetectError = '';
        st.detectedOnce = false;
        delete st._detectToken;
        fileUiCache[fid] = st;
        var rrow = __batchWorkbenchRows[String(fid)];
        if (rrow) {
          rrow._wbPipelineDone = false;
          rrow._wbMatchAttempted = false;
          rrow.manualDetectWrong = false;
          rrow.status = '识别中…';
          rrow.slotTags = ['识别中'];
        }
        try { renderBatchWorkbenchTable(); } catch (_) {}
        withButtonBusy(btnDet, '识别中…', function () {
          return processBatchWorkbenchFileIds([fid], true, {
            hideMask: true,
            forceReprocess: true,
          });
        }, { skipPageProgress: true });
      });
      var btnCanvas = document.createElement('button');
      btnCanvas.type = 'button';
      btnCanvas.className = 'btn btn-secondary';
      btnCanvas.style.padding = '4px 8px';
      btnCanvas.style.fontSize = '0.82rem';
      btnCanvas.style.marginLeft = '4px';
      btnCanvas.textContent = '画布';
      btnCanvas.title = '高级模式：为未匹配库素材的角色手写补签';
      btnCanvas.addEventListener('click', function () {
        selectWorkbenchFileForAdvanced(fid);
      });
      var btnWrong = document.createElement('button');
      btnWrong.type = 'button';
      btnWrong.className = 'btn btn-secondary';
      btnWrong.style.padding = '4px 8px';
      btnWrong.style.fontSize = '0.82rem';
      btnWrong.style.marginLeft = '4px';
      btnWrong.textContent = row.manualDetectWrong || _stShown === '识别有误' ? '改登记' : '标误';
      btnWrong.title =
        '登记识别纠正：勾选保存需签角色和/或签字位版式，重新识别时自动带入';
      btnWrong.addEventListener('click', function (ev) {
        ev.stopPropagation();
        openDetectCorrectionDialog(fid);
      });
      tdOp.appendChild(btnWrong);
      tdOp.appendChild(btnDet);
      tdOp.appendChild(btnCanvas);

      tr.addEventListener('click', function (ev) {
        if (ev.target.closest('button,input,select,label,a')) return;
        selectedFileId = fid;
        __batchWorkbenchEditFileId = fid;
      });

      tr.appendChild(td0);
      tr.appendChild(tdName);
      tr.appendChild(tdDocDate);
      tr.appendChild(tdLoc);
      tr.appendChild(tdEdSig);
      tr.appendChild(tdEdDate);
      tr.appendChild(tdRvSig);
      tr.appendChild(tdRvDate);
      tr.appendChild(tdApSig);
      tr.appendChild(tdApDate);
      tr.appendChild(tdMat);
      tr.appendChild(tdRoles);
      tr.appendChild(tdSlot);
      tr.appendChild(tdSt);
      tr.appendChild(tdOp);
      batchWorkbenchBody.appendChild(tr);
    });
    var visRows = [];
    savedFiles.forEach(function (r) {
      if (!r || !r.id) return;
      var rw = __batchWorkbenchRows[String(r.id)];
      if (!rw || !_workbenchRowPassesFilter(r, rw)) return;
      visRows.push(rw);
    });
    var allOn = visRows.length > 0 && visRows.every(function (rw) {
      return rw && rw.selected;
    });
    if (batchWorkbenchHeadCheck) batchWorkbenchHeadCheck.checked = allOn;
    if (batchWorkbenchSelectAll) {
      batchWorkbenchSelectAll.checked = _workbenchHasActiveFilter()
        ? allOn
        : savedFiles.length > 0 &&
          savedFiles.every(function (r) {
            var rw = __batchWorkbenchRows[String(r.id)];
            return rw && rw.selected;
          });
    }
    _updateWorkbenchFilterHint(shown, total);
  }

  function processBatchWorkbenchFile(fileId, opts) {
    opts = opts && typeof opts === 'object' ? opts : {};
    var row = __batchWorkbenchRows[String(fileId)];
    if (!row) return Promise.resolve();
    if (row._wbPipelineDone && !opts.forceReprocess) {
      return Promise.resolve();
    }
    if (shouldSkipDetectPipelineForFile(fileId) && !opts.forceReprocess) {
      row.status = '无需签字';
      row.rolesLabel = '无需签字（用例表）';
      row.slotLabel = '无需签字';
      row.slotExplain = '规则标记该文档无需签字。';
      row.slotTags = ['无需签字'];
      row.slotSignable = true;
      row.detectedRoleIds = [];
      row._wbPipelineDone = true;
      row._wbMatchAttempted = true;
      syncWorkbenchRowFromDetect(fileId);
      schedulePersistFileSessionCache(fileId);
      renderBatchWorkbenchTable();
      return Promise.resolve();
    }
    var skipDetect = !!opts.skipDetectIfCached && fileHasValidDetectCache(fileId);
    syncRowToHandoffCtx(fileId);
    var prevSel = selectedFileId;
    var ctx = rowCtxFromRowState(row);
    __aiwordHandoffCtxByFileId[String(fileId)] = ctx;
    __aiwordHandoffCtx = ctx;
    __aiwordHandoffTargetFileId = fileId;
    selectedFileId = fileId;
    if (!fileUiCache[fileId]) fileUiCache[fileId] = {};
    row._wbProcessing = true;
    if (!skipDetect) {
      row.status = '识别中…';
      row.slotTags = ['识别中'];
      fileUiCache[fileId].lastDetectData = null;
      fileUiCache[fileId].lastDetectError = '';
      fileUiCache[fileId].detectedOnce = false;
      syncWorkbenchRowFromDetect(fileId, null);
    } else {
      row.status = '匹配素材…';
    }
    renderBatchWorkbenchTable();

    // signersList 由 processBatchWorkbenchFileIds 在批量入口已统一刷新一次，
    // 这里只在签署人库尚未就绪时兜底，避免单次抖动让后续所有文件「找不到签署人」。
    var ensureSigners = signersList && signersList.length
      ? Promise.resolve()
      : refreshSigners();
    return ensureSigners
      .then(function () {
        return ensureFileRoleMapLoaded(fileId);
      })
      .then(function () {
        currentRoleMap = _deepCloneJsonish((fileUiCache[fileId] || {}).currentRoleMap || {});
        _lastPersistedRoleMapStable[String(fileId)] = _roleMapStableJson(
          (fileUiCache[fileId] || {}).currentRoleMap || {}
        );
        if (skipDetect) return true;
        function _detectOnceOrRetryLeft(retryLeft) {
          var detectP = detectAndAutoSelectRoles(fileId, __aiwordHandoffDetectRetries, {
            returnPromise: true,
            batchSilent: true,
          });
          return _withBatchDetectTimeout(detectP, row).then(function (ok) {
            if (ok || retryLeft <= 0) return !!ok;
            var waitMs = 800 + (2 - retryLeft) * 700;
            row.status = '识别重试中…';
            row.slotTags = ['识别中'];
            renderBatchWorkbenchTable();
            return new Promise(function (resolve) {
              setTimeout(resolve, waitMs);
            }).then(function () {
              if (!fileUiCache[fileId]) fileUiCache[fileId] = {};
              fileUiCache[fileId].detectedOnce = false;
              return _detectOnceOrRetryLeft(retryLeft - 1);
            });
          });
        }
        return _detectOnceOrRetryLeft(2);
      })
      .then(function (detectOk) {
        syncWorkbenchRowFromDetect(fileId);
        var stDet = (fileUiCache[fileId] || {}).lastDetectData;
        var roleList = mergeDetectedRolesForFile(fileId);
        if (detectOk === false || !stDet || !stDet.ok) {
          row.status = '识别失败';
          row.slotTags = ['识别失败', '不可签'];
          return Promise.reject(new Error('识别失败'));
        }
        if (!roleList.length) {
          row.status = '未识别到签字位';
          row.rolesLabel =
            (fileUiCache[fileId] || {}).lastDetectError ||
            '未匹配到编制/编写/审核/批准等标签';
          row.slotTags = ['未识别到签字位', '不可签'];
          setBatchWorkbenchMsg(
            (row.name || fileId) + '：' + row.rolesLabel + '（可检查 Word 签批栏是否在页脚表格）',
            true
          );
          return Promise.reject(new Error('未识别到签字位'));
        }
        if (opts.detectOnly) {
          finalizeWorkbenchRowStatus(fileId);
          return null;
        }
        row.status = '匹配素材…';
        renderBatchWorkbenchTable();
        return applyAiwordHandoffHintsOnce(fileId, ctx, {
          skipSelectedCheck: true,
          batchSilent: true,
        });
      })
      .then(function () {
        if (opts.detectOnly) return;
        finalizeWorkbenchRowStatus(fileId);
      })
      .catch(function (e) {
        if (shouldSkipDetectPipelineForFile(fileId) || fileHasValidDetectCache(fileId)) {
          finalizeWorkbenchRowStatus(fileId);
        } else {
          var msg = (e && e.message) || '识别失败';
          var nm = row.name || fileId;
          if (/超时|timeout|aborted|已取消上传/i.test(String(msg))) {
            row.status = '识别超时';
            row.rolesLabel = '识别超时（' + msg + '）。可在右侧素材区手动选择签名后再次保存，或在「系统设置 → SIGN_DETECT_TIMEOUT_MS」中加大超时。';
            row.slotTags = ['识别超时', '不可签'];
            setBatchWorkbenchMsg(
              '「' + nm + '」识别超时（' + msg + '），将继续处理下一个文件。\n' +
                '该文件可稍后在表格点击「识别」按钮单独重试，' +
                '或在「系统设置」中加大 SIGN_DETECT_TIMEOUT_MS 后重试。',
              'warn'
            );
          } else {
            row.status = '识别失败';
            row.slotTags = ['识别失败', '不可签'];
            setBatchWorkbenchMsg(
              '「' + nm + '」识别失败：' + msg + '，将继续处理下一个文件。',
              'warn'
            );
          }
        }
        // 识别失败/超时时仍按编审批姓名匹配一次（in-flight 合并不会重复 PUT）
        if (!opts.detectOnly && !row._wbMatchAttempted) {
          return workbenchSyncUiAfterMatch(fileId).catch(function () {});
        }
      })
      .then(function () {
        row._wbPipelineDone = true;
        saveCurrentFileUiToCache(fileId);
        selectedFileId = prevSel;
        renderBatchWorkbenchTable();
        syncHiddenBatchPicks();
      })
      .finally(function () {
        row._wbProcessing = false;
      });
  }

  function processBatchWorkbenchFileIds(fileIds, detectOnly, runOpt) {
    runOpt = runOpt && typeof runOpt === 'object' ? runOpt : {};
    var ids = (fileIds || []).filter(Boolean);
    if (!ids.length) {
      setBatchWorkbenchMsg('没有可处理的文件', true);
      setBatchWorkbenchProgress(0, 0, '', false);
      return Promise.resolve();
    }
    var wbMode = isBatchWorkbenchMode();
    if (!runOpt.skipPageProgress && !wbMode) {
      beginPageProgress(
        detectOnly ? '批量识别角色与签字位…' : '批量识别并匹配素材…',
        { done: 0, total: ids.length }
      );
    }
    if (!wbMode) {
      setBatchWorkbenchMsg('正在处理 ' + ids.length + ' 个文件…', false);
    }
    setBatchWorkbenchProgress(
      0,
      ids.length,
      (detectOnly ? '批量识别（角色+签字位）' : '识别+匹配') +
        ' 0/' +
        ids.length +
        '（0%）',
      true
    );
    // 批量入口统一刷新一次签署人库；后续每文件不再单独触发刷新，
    // 防止某次抖动让 signersList 清空、后续文件全都「找不到签署人」。
    // 若 boot 阶段已经拉过且 signersList 非空，则直接跳过这次刷新，避免重复请求拖慢启动。
    var chain = signersList && signersList.length
      ? Promise.resolve()
      : refreshSigners().catch(function () {});
    var startTs = Date.now();
    ids.forEach(function (fid, idx) {
      chain = chain.then(function () {
        var row = __batchWorkbenchRows[String(fid)];
        var nm = row && row.name ? row.name : fid;
        var nmShort = nm.length > 28 ? nm.slice(0, 26) + '…' : nm;
        var elapsedSec = Math.max(0, Math.round((Date.now() - startTs) / 1000));
        var curPct = ids.length ? Math.round((idx / ids.length) * 100) : 0;
        setBatchWorkbenchProgress(
          idx,
          ids.length,
          (detectOnly ? '批量识别（角色+签字位）' : '识别+匹配') +
            ' ' +
            idx +
            '/' +
            ids.length +
            '（' +
            curPct +
            '%） · ' +
            nmShort +
            ' · 已耗时 ' +
            elapsedSec +
            's',
          true
        );
        return processBatchWorkbenchFile(fid, {
          detectOnly: !!detectOnly,
          skipDetectIfCached: !!runOpt.skipDetectIfCached,
          forceReprocess: !!runOpt.forceReprocess,
        });
      });
    });
    return chain
      .then(function () {
        return flushPendingFileSessionCaches(ids);
      })
      .then(function () {
      var totalSec = Math.max(0, Math.round((Date.now() - startTs) / 1000));
      var sum = _summarizeBatchWorkbenchProgress(ids);
      var failN = sum.detectTimeout + sum.detectFail + sum.noField;
      var head = '已完成 ' + ids.length + ' 个文件' + (detectOnly ? '识别' : '识别+匹配') +
        '（耗时 ' + totalSec + 's）。';
      var detail = '就绪 ' + sum.ready + ' / 部分就绪 ' + sum.partial +
        ' / 待匹配 ' + sum.waitMatch +
        (sum.detectWrong ? ' / 识别有误 ' + sum.detectWrong : '') +
        (sum.noSign ? ' / 无需签字 ' + sum.noSign : '') +
        ' / 识别超时 ' + sum.detectTimeout +
        ' / 识别失败 ' + sum.detectFail +
        ' / 未识别到签字位 ' + sum.noField + '。';
      var lines = [head, detail];
      if (sum.failItems.length) {
        lines.push('失败明细（共 ' + failN + ' 个）：');
        sum.failItems.slice(0, 8).forEach(function (it) {
          lines.push(' • [' + it.status + '] ' + it.name +
            (it.detail ? '：' + it.detail : ''));
        });
        if (sum.failItems.length > 8) {
          lines.push(' …其余 ' + (sum.failItems.length - 8) + ' 项可在表格状态列查看。');
        }
        if (sum.detectTimeout) {
          lines.push('提示：识别超时通常是「文档很大 + MySQL/FTP 拉取慢」。' +
            '可在「系统设置」中加大 SIGN_DETECT_TIMEOUT_MS，或直接在右侧素材区手动选择签名后保存。');
        }
      }
      setBatchWorkbenchMsg(lines.join('\n'), failN ? 'warn' : 'ok');
      setBatchWorkbenchProgress(
        ids.length,
        ids.length,
        (detectOnly ? '批量识别（角色+签字位）' : '识别+匹配') +
          ' ' +
          ids.length +
          '/' +
          ids.length +
          '（100%） · 已完成',
        true
      );
      setTimeout(function () {
        setBatchWorkbenchProgress(0, 0, '', false);
      }, 5000);
      if (runOpt.hideMask !== false) {
        _hideAiwordHandoffLoadingMask();
      }
      if (!detectOnly && ids.length && isBatchWorkbenchMode()) {
        var pick =
          __batchWorkbenchEditFileId ||
          _firstAiwordHandoffFileId() ||
          ids[0];
        selectWorkbenchFileForMaterialEdit(pick, { scroll: false });
      }
      try {
        _refreshWorkbenchSlotFilterOptions();
      } catch (_) {}
    })
    .finally(function () {
      if (!runOpt.skipPageProgress) endPageProgress();
    });
  }

  function processBatchWorkbenchAllFiles(detectOnly) {
    var ids = savedFiles
      .map(function (r) {
        return r && r.id;
      })
      .filter(Boolean);
    return processBatchWorkbenchFileIds(ids, detectOnly, { hideMask: false });
  }

  function processBatchWorkbenchSelected(detectOnly, runOpt) {
    runOpt = runOpt && typeof runOpt === 'object' ? runOpt : {};
    var ids = savedFiles
      .map(function (r) {
        return r && r.id;
      })
      .filter(function (id) {
        var row = __batchWorkbenchRows[String(id)];
        return row && row.selected;
      });
    if (!ids.length) {
      setBatchWorkbenchMsg('请先勾选要处理的文件（可先「全选当前筛选」）', true);
      return Promise.resolve();
    }
    ids.forEach(function (fid) {
      var row = __batchWorkbenchRows[String(fid)];
      if (row && runOpt.forceReprocess) {
        row._wbPipelineDone = false;
        row._wbMatchAttempted = false;
        row.manualDetectWrong = false;
        row.detectWrongNote = '';
      }
    });
    return processBatchWorkbenchFileIds(ids, detectOnly, runOpt);
  }

  var __batchWorkbenchPipelineRunning = false;
  var _wbAutoMatchTimer = null;
  var _wbAutoMatchInFlight = false;

  function onAiwordHandoffFilesReady() {
    var n = Object.keys(__aiwordHandoffCtxByFileId || {}).length;
    if ((!_isFromAiwordHandoff() && !__aiwordPendingBatchWorkbench) || n < 1) {
      return Promise.resolve();
    }
    if (__wbHandoffPipelineStarted) {
      return Promise.resolve();
    }
    __wbHandoffPipelineStarted = true;
    __aiwordPendingBatchWorkbench = false;
    enableBatchWorkbenchMode();
    return refreshWorkbenchFilesFromServer({ silent: true }).then(function () {
    initBatchRowsFromHandoff();
    renderBatchWorkbenchTable();
    syncHiddenBatchPicks();
    if (batchSelectAll) batchSelectAll.checked = false;
    if (batchWorkbenchSelectAll) batchWorkbenchSelectAll.checked = false;
    if (batchWorkbenchHeadCheck) batchWorkbenchHeadCheck.checked = false;
    setBatchWorkbenchMsg(
      n <= 5
        ? '已载入 ' + n + ' 个 aiword 任务，正在后台识别并匹配（约 ' + n * 5 + '~' + n * 15 + ' 秒）…'
        : '已载入 aiword 任务 ' +
            n +
            ' 个（列表中其它历史文件未默认勾选），正在后台识别并匹配…',
      false
    );
    // claim 已完成、表格已渲染：立刻关闭 mask 让用户能立即看到表格并操作；
    // 识别/匹配在后台跑，进度通过 setBatchWorkbenchMsg 与每行 status 列实时反馈，
    // 避免出现「页面提示 1 分钟、实际更长」的等待错觉。
    _hideAiwordHandoffLoadingMask();
    var ids = _aiwordHandoffFileIdSet();
    __batchWorkbenchPipelineRunning = true;
    // refreshSigners 由 processBatchWorkbenchFileIds 内部统一发起（且支持「已加载就跳过」）
    return processBatchWorkbenchFileIds(ids, false, { hideMask: true })
      .then(function () {
        var pick = _firstAiwordHandoffFileId();
        if (pick) {
          selectWorkbenchFileForMaterialEdit(pick, { scroll: false });
        }
      })
      .catch(function (e) {
        setBatchWorkbenchMsg(
          '批量处理出错：' + ((e && e.message) || '未知错误') +
            '。可在表格中单独点击每行「识别」按钮重试，或刷新页面后再次进入。',
          'error'
        );
      })
      .finally(function () {
        __batchWorkbenchPipelineRunning = false;
        // 成功路径已在 processBatchWorkbenchFileIds 末尾给出汇总；
        // 这里仅在汇总为空（异常退出）时兜底显示结束消息。
        if (batchWorkbenchMsg && !batchWorkbenchMsg.textContent) {
          var sum = _summarizeBatchWorkbenchProgress(ids);
          setBatchWorkbenchMsg(
            '后台识别匹配已结束。就绪 ' + sum.ready +
              ' / 部分就绪 ' + sum.partial +
              ' / 待匹配 ' + sum.waitMatch +
        (sum.detectWrong ? ' / 识别有误 ' + sum.detectWrong : '') +
              (sum.noSign ? ' / 无需签字 ' + sum.noSign : '') +
              ' / 识别超时 ' + sum.detectTimeout +
              ' / 识别失败 ' + sum.detectFail +
              ' / 未识别到签字位 ' + sum.noField + '。',
            (sum.detectTimeout + sum.detectFail + sum.noField) ? 'warn' : 'ok'
          );
        }
      });
    });
  }

  function startManualUploadBatchWorkbench(fileIds, opts) {
    opts = opts && typeof opts === 'object' ? opts : {};
    var ids = (fileIds || []).map(String).filter(Boolean);
    if (!ids.length) return Promise.resolve();
    enableBatchWorkbenchMode();
    var afterRefresh = opts.skipFileListRefresh
      ? Promise.resolve()
      : refreshWorkbenchFilesFromServer({ silent: true });
    return afterRefresh.then(function () {
    ids.forEach(function (fid) {
      if (!__aiwordHandoffCtxByFileId[fid] || typeof __aiwordHandoffCtxByFileId[fid] !== 'object') {
        __aiwordHandoffCtxByFileId[fid] = {};
      }
      if (__batchWorkbenchRows[fid]) {
        __batchWorkbenchRows[fid].selected = true;
        if (!__batchWorkbenchRows[fid].status || __batchWorkbenchRows[fid].status === '—') {
          __batchWorkbenchRows[fid].status = '待处理';
        }
      }
    });
    renderBatchWorkbenchTable();
    syncHiddenBatchPicks();
    if (batchSelectAll) batchSelectAll.checked = false;
    if (batchWorkbenchSelectAll) batchWorkbenchSelectAll.checked = false;
    if (batchWorkbenchHeadCheck) batchWorkbenchHeadCheck.checked = false;
    setBatchWorkbenchMsg(
      '已载入 ' + ids.length + ' 个手动上传文件，正在后台识别签字位并匹配素材…',
      false
    );
    if (__batchWorkbenchPipelineRunning) {
      setBatchWorkbenchMsg(
        '上一轮批量识别仍在进行，新增文件已加入列表，可稍后点击「批量识别（角色+签字位）」继续。',
        'warn'
      );
      return Promise.resolve();
    }
    __batchWorkbenchPipelineRunning = true;
    return processBatchWorkbenchFileIds(ids, false, {
      hideMask: true,
      skipDetectIfCached: !!opts.skipDetectIfCached,
    })
      .then(function () {
        var pick = ids[0];
        if (pick) selectWorkbenchFileForMaterialEdit(pick, { scroll: false });
      })
      .catch(function (e) {
        setBatchWorkbenchMsg(
          '手动上传批量处理出错：' + ((e && e.message) || '未知错误') + '。可在表格中按行重试。',
          'error'
        );
      })
      .finally(function () {
        __batchWorkbenchPipelineRunning = false;
      });
    });
  }

  function renderFileList() {
    if (!IS_FILE_SIGN_PAGE || !fileListEl || !listHint || !needSignTable) return;
    if (savedFiles.length) {
      enableBatchWorkbenchMode();
    }
    if (isBatchWorkbenchMode()) {
      syncHiddenBatchPicks();
      renderBatchWorkbenchTable();
      updateSubmitState();
      updateBatchUi();
      return;
    }
    fileListEl.innerHTML = '';
    if (!savedFiles.length) {
      listHint.style.display = 'block';
      listHint.textContent = '暂无已保存文件，请先上传保存。';
      selectedFileId = null;
      lastDetectData = null;
      lastDetectFileId = null;
      lastDetectError = '';
      detectEpoch += 1;
      detectInFlightFor = null;
      currentRoleMap = {};
      needSignTable.innerHTML = '';
      resetAllRoleChecks();
      updateSubmitState();
      updateBatchUi();
      return;
    }
    listHint.style.display = 'none';
    savedFiles.forEach(function (rec) {
      var li = document.createElement('li');
      if (rec.id === selectedFileId) {
        li.classList.add('selected');
      }
      var batchCb = document.createElement('input');
      batchCb.type = 'checkbox';
      batchCb.className = 'batch-pick';
      batchCb.setAttribute('data-id', rec.id);
      batchCb.title = '加入批量签名';
      var wbRow = __batchWorkbenchRows[String(rec.id)];
      batchCb.checked = wbRow ? !!wbRow.selected : _isAiwordHandoffFileId(rec.id);
      batchCb.addEventListener('change', updateBatchUi);
      var rid = 'pick_' + rec.id;
      var radio = document.createElement('input');
      radio.type = 'radio';
      radio.name = 'savedFile';
      radio.id = rid;
      radio.value = rec.id;
      if (rec.id === selectedFileId) {
        radio.checked = true;
      }
      var lbl = document.createElement('label');
      lbl.htmlFor = rid;
      var ftpHint =
        typeof rec.ftp_uploaded === 'boolean'
          ? rec.ftp_uploaded
            ? ' [FTP 已上传]'
            : ' [FTP 未上传，内容在库内' +
              (rec.ftp_last_error
                ? '：' +
                  String(rec.ftp_last_error).slice(0, 48) +
                  (String(rec.ftp_last_error).length > 48 ? '…' : '')
                : '') +
              ']'
          : '';
      lbl.textContent = (rec.name || rec.id) + ftpHint;
      var delBtn = document.createElement('button');
      delBtn.type = 'button';
      delBtn.className = 'btn btn-secondary del-btn';
      delBtn.textContent = '删除';

      radio.addEventListener('change', function () {
        // 切换前：保存当前文件的 UI 状态（已选角色/筛选输入等），避免切回丢失
        saveCurrentFileUiToCache(selectedFileId);
        selectedFileId = rec.id;
        document.querySelectorAll('.file-list li').forEach(function (el) {
          el.classList.remove('selected');
        });
        li.classList.add('selected');
        // 切换后：若之前已识别过该文件，则直接恢复，不再自动识别；首次切换才自动识别
        if (!restoreFileUiFromCache(selectedFileId)) {
          var st0 = fileUiCache[selectedFileId] || {};
          if (!st0.detectedOnce) {
            detectAndAutoSelectRoles(selectedFileId);
          } else {
            // 仅加载 role-map，确保表格可编辑
            fetchJson(apiUrl('/api/sign/files/' + selectedFileId + '/role-map'))
              .then(function (r) {
                var jj = r.data || {};
                if (jj.ok) currentRoleMap = jj.map || {};
                cachePatchCurrentRoleMap(selectedFileId, currentRoleMap);
                renderNeedSignTable();
                updateSubmitState();
              })
              .catch(function () {});
          }
        } else {
          // 已恢复：仍从服务端刷新一次 role-map（防止多端修改）
          fetchJson(apiUrl('/api/sign/files/' + selectedFileId + '/role-map'))
            .then(function (r) {
              var jj = r.data || {};
              if (jj.ok) currentRoleMap = jj.map || {};
              cachePatchCurrentRoleMap(selectedFileId, currentRoleMap);
              renderNeedSignTable();
              updateSubmitState();
            })
            .catch(function () {});
        }
        // 若该文件属于“批量映射目标”，才显示提示；否则清空提示
        showNeedSignNoticeForSelectedFile();
      });
      delBtn.addEventListener('click', function () {
        withButtonBusy(delBtn, '删除中…', function () {
          return fetchJson(apiUrl('/api/sign/files/' + rec.id), { method: 'DELETE' }).then(function (
            result
          ) {
            var j = result.data;
            if (!j.ok) {
              setFileListActionFeedback(j.error || '删除失败', true);
              return;
            }
            setFileListActionFeedback('', false);
            if (
              __aiwordHandoffTargetFileId != null &&
              String(__aiwordHandoffTargetFileId) === String(rec.id)
            ) {
              clearAiwordHandoffState();
            }
            if (selectedFileId === rec.id) {
              selectedFileId = null;
            }
            return refreshFileList().catch(function () {
              savedFiles = j.files || [];
              renderFileList();
            });
          });
        }).catch(function (e) {
          setFileListActionFeedback(e.message || String(e), true);
        });
      });

      li.appendChild(batchCb);
      li.appendChild(radio);
      li.appendChild(lbl);
      li.appendChild(delBtn);
      fileListEl.appendChild(li);
    });
    if (!selectedFileId && savedFiles.length) {
      var autoPick = _aiwordHandoffFileIdSet().length
        ? _firstAiwordHandoffFileId()
        : savedFiles[0].id;
      if (autoPick) {
        selectedFileId = autoPick;
        renderFileList();
        return;
      }
    }
    updateSubmitState();
    updateBatchUi();
    var sid = selectedFileId;
    // 首次进入页面：仅首次选中时自动识别；若曾识别过（缓存里有标记），不再自动识别
    if (sid) {
      var st = fileUiCache[sid] || {};
      var deferAi =
        typeof __aiwordDeferDetectFileId !== 'undefined' &&
        __aiwordDeferDetectFileId != null &&
        String(__aiwordDeferDetectFileId) === String(sid);
      if (__aiwordPendingBatchWorkbench || isBatchWorkbenchMode()) {
        __aiwordDeferDetectFileId = null;
      } else if (deferAi) {
        __aiwordDeferDetectFileId = null;
        setTimeout(function () {
          if (String(selectedFileId) !== String(sid)) return;
          kickoffAiwordHandoffPipeline(sid);
        }, 520);
      } else if (!st.detectedOnce) {
        detectAndAutoSelectRoles(sid);
      } else {
        // 有缓存则恢复；并刷新 role-map
        restoreFileUiFromCache(sid);
        fetchJson(apiUrl('/api/sign/files/' + sid + '/role-map'))
          .then(function (r) {
            var jj = r.data || {};
            if (jj.ok) currentRoleMap = jj.map || {};
            cachePatchCurrentRoleMap(sid, currentRoleMap);
            renderNeedSignTable();
          })
          .catch(function () {});
      }
    }
  }

  function formatRolesLabel(rolesJson) {
    if (!rolesJson) return '';
    try {
      var ids = JSON.parse(rolesJson);
      if (!Array.isArray(ids)) return '';
      return ids
        .map(function (id) {
          var x = ROLES.find(function (r) {
            return r.id === id;
          });
          return x ? x.label : id;
        })
        .join('、');
    } catch (_) {
      return '';
    }
  }

  function renderSignedList(j) {
    if (!IS_FILE_SIGN_PAGE || !signedListEl || !signedHint) return;
    var batches = (j && j.batches) || [];
    var items = (j && j.items) || [];
    var dbShare = !!(j && j.db_share);
    var total = typeof (j && j.total) === 'number' ? j.total : (batches.length ? batches.length : items.length);
    var page = typeof (j && j.page) === 'number' ? j.page : 1;
    var pageSize = typeof (j && j.page_size) === 'number' ? j.page_size : signedListPageSize;
    var pageCountEarly = Math.max(1, Math.ceil(total / pageSize) || 1);
    if (dbShare && total > 0 && page > pageCountEarly) {
      signedListPage = pageCountEarly;
      refreshSignedList();
      return;
    }
    signedListEl.innerHTML = '';
    signedSearchRow.style.display = 'none';
    if (!dbShare) {
      signedHint.style.display = 'block';
      signedHint.textContent =
        '当前未配置 MySQL（环境变量 MYSQL_HOST）。配置并重启服务后，生成成功的已签名文件会写入数据库，局域网内其他电脑打开本页即可从下列表下载。';
      return;
    }
    signedSearchRow.style.display = 'block';
    if (!batches.length && !items.length) {
      signedHint.style.display = 'block';
      signedHint.textContent = signedListQ
        ? '无匹配的已签名记录，请尝试其它关键字或清空搜索。'
        : '暂无已签名记录。点击「生成已签名文档」成功后，文件会保存到数据库并出现在此列表。';
    } else {
      signedHint.style.display = 'none';
    }
    var pageCount = Math.max(1, Math.ceil(total / pageSize) || 1);
    signedPagerInfo.textContent =
      '共 ' +
      total +
      ' 条 · 第 ' +
      page +
      ' / ' +
      pageCount +
      ' 页（每页 ' +
      pageSize +
      ' 条）';
    signedPrevBtn.disabled = page <= 1;
    signedNextBtn.disabled = page >= pageCount;
    function renderSignedItemRow(it, parent) {
      var li = document.createElement('li');
      var wrap = document.createElement('div');
      wrap.style.flex = '1';
      wrap.style.minWidth = '0';
      var title = document.createElement('div');
      title.style.fontWeight = '500';
      title.style.wordBreak = 'break-all';
      title.textContent = it.name || it.id;
      var meta = document.createElement('div');
      meta.className = 'signed-meta';
      var parts = [];
      if (it.created_at) parts.push(it.created_at);
      var rl = formatRolesLabel(it.roles_json);
      if (rl) parts.push('签字角色：' + rl);
      var signedFtp = ftpMetaLine(
        it.ftp_uploaded,
        it.ftp_last_error,
        '已保存在 MySQL，可下载；可稍后重试迁移'
      );
      if (signedFtp) parts.push(signedFtp);
      meta.textContent = parts.join(' · ');
      wrap.appendChild(title);
      wrap.appendChild(meta);

      var dl = document.createElement('a');
      dl.className = 'btn btn-secondary';
      dl.href = apiUrl('/api/sign/signed/' + it.id);
      dl.setAttribute('download', '');
      dl.textContent = '下载';

      var delBtn = document.createElement('button');
      delBtn.type = 'button';
      delBtn.className = 'btn btn-secondary del-btn';
      delBtn.textContent = '删除';
      delBtn.addEventListener('click', function () {
        withButtonBusy(delBtn, '删除中…', function () {
          return fetchJson(apiUrl('/api/sign/signed/' + it.id), { method: 'DELETE' }).then(
            function (result) {
              var jj = result.data;
              if (!jj.ok) {
                setSignedListActionFeedback(jj.error || '删除失败', true);
                return;
              }
              setSignedListActionFeedback('', false);
              refreshSignedList();
            }
          );
        }).catch(function (e) {
          setSignedListActionFeedback(e.message || String(e), true);
        });
      });

      li.appendChild(wrap);
      li.appendChild(dl);
      li.appendChild(delBtn);
      parent.appendChild(li);
    }

    // 历史记录（未分批）分组：旧数据 batch_id 为空（即使当前没有任何新批次，也要展示）
    var legacyTotal = typeof (j && j.legacy_total) === 'number' ? j.legacy_total : 0;
    if (legacyTotal > 0) {
      (function () {
        var li = document.createElement('li');
        var wrap = document.createElement('div');
        wrap.style.flex = '1';
        wrap.style.minWidth = '0';
        var title = document.createElement('div');
        title.style.fontWeight = '600';
        title.style.wordBreak = 'break-all';
        title.textContent = '历史记录（未分批）（' + legacyTotal + ' 个）';
        var meta = document.createElement('div');
        meta.className = 'signed-meta';
        meta.textContent = '旧版生成的已签名记录（无批次号）。可展开查看/搜索/下载 zip。';
        wrap.appendChild(title);
        wrap.appendChild(meta);

        var toggle = document.createElement('button');
        toggle.type = 'button';
        toggle.className = 'btn btn-secondary';
        toggle.textContent = '展开';

        var zipA = document.createElement('a');
        zipA.className = 'btn btn-secondary';
        zipA.textContent = '下载历史包';
        zipA.href = apiUrl('/api/sign/signed-legacy/zip?q=' + encodeURIComponent(signedListQ || ''));
        zipA.setAttribute('download', '');

        var box = document.createElement('div');
        box.style.display = 'none';
        box.style.width = '100%';
        box.style.marginTop = '10px';
        box.style.paddingTop = '10px';
        box.style.borderTop = '1px solid var(--border)';

        var searchRow = document.createElement('div');
        searchRow.className = 'btn-row';
        searchRow.style.margin = '0 0 10px';
        var bi = document.createElement('input');
        bi.type = 'search';
        bi.placeholder = '在历史记录内搜索文件名…';
        bi.style.flex = '1';
        bi.style.minWidth = '140px';
        bi.style.padding = '8px 12px';
        bi.style.border = '1px solid var(--border)';
        bi.style.borderRadius = '8px';
        var bs = document.createElement('button');
        bs.type = 'button';
        bs.className = 'btn btn-secondary';
        bs.textContent = '搜索';
        searchRow.appendChild(bi);
        searchRow.appendChild(bs);

        var pager = document.createElement('div');
        pager.className = 'btn-row';
        pager.style.margin = '0 0 10px';
        pager.style.alignItems = 'center';
        var pi = document.createElement('span');
        pi.className = 'hint';
        pi.style.margin = '0';
        var pprev = document.createElement('button');
        pprev.type = 'button';
        pprev.className = 'btn btn-secondary';
        pprev.textContent = '上一页';
        var pnext = document.createElement('button');
        pnext.type = 'button';
        pnext.className = 'btn btn-secondary';
        pnext.textContent = '下一页';
        pager.appendChild(pi);
        pager.appendChild(pprev);
        pager.appendChild(pnext);

        var innerList = document.createElement('ul');
        innerList.className = 'file-list';
        innerList.style.marginTop = '0';

        var legacyPage = 1;
        var legacyPageSize = 20;
        var legacyQ = '';

        function loadLegacy() {
          innerList.innerHTML = '';
          var h = document.createElement('li');
          h.textContent = '正在加载历史记录…';
          h.style.border = 'none';
          h.style.background = 'transparent';
          innerList.appendChild(h);
          fetchJson(
            apiUrl(
              '/api/sign/signed-legacy?q=' +
                encodeURIComponent(legacyQ || '') +
                '&page=' +
                legacyPage +
                '&page_size=' +
                legacyPageSize +
                '&_=' +
                Date.now()
            )
          )
            .then(function (r) {
              var jj = r.data || {};
              if (!jj.ok) {
                innerList.innerHTML = '';
                var e = document.createElement('li');
                e.textContent = '加载失败：' + (jj.error || '未知错误');
                e.style.border = 'none';
                e.style.background = 'transparent';
                innerList.appendChild(e);
                return;
              }
              var arr = jj.items || [];
              var total2 = typeof jj.total === 'number' ? jj.total : arr.length;
              var page2 = typeof jj.page === 'number' ? jj.page : legacyPage;
              var ps2 = typeof jj.page_size === 'number' ? jj.page_size : legacyPageSize;
              var pc = Math.max(1, Math.ceil(total2 / ps2) || 1);
              pi.textContent = '共 ' + total2 + ' 条 · 第 ' + page2 + ' / ' + pc + ' 页';
              pprev.disabled = page2 <= 1;
              pnext.disabled = page2 >= pc;
              innerList.innerHTML = '';
              if (!arr.length) {
                var e2 = document.createElement('li');
                e2.textContent = legacyQ ? '无匹配文件。' : '暂无历史记录。';
                e2.style.border = 'none';
                e2.style.background = 'transparent';
                innerList.appendChild(e2);
                return;
              }
              arr.forEach(function (it) {
                renderSignedItemRow(it, innerList);
              });
            })
            .catch(function (e) {
              innerList.innerHTML = '';
              var e3 = document.createElement('li');
              e3.textContent = '加载失败：' + (e && e.message ? e.message : String(e));
              e3.style.border = 'none';
              e3.style.background = 'transparent';
              innerList.appendChild(e3);
            });
        }

        bs.addEventListener('click', function () {
          legacyQ = bi.value ? String(bi.value).trim() : '';
          legacyPage = 1;
          zipA.href = apiUrl('/api/sign/signed-legacy/zip?q=' + encodeURIComponent(legacyQ || ''));
          loadLegacy();
        });
        bi.addEventListener('keydown', function (ev) {
          if (ev && ev.key === 'Enter') {
            ev.preventDefault();
            legacyQ = bi.value ? String(bi.value).trim() : '';
            legacyPage = 1;
            zipA.href = apiUrl('/api/sign/signed-legacy/zip?q=' + encodeURIComponent(legacyQ || ''));
            loadLegacy();
          }
        });
        pprev.addEventListener('click', function () {
          if (legacyPage <= 1) return;
          legacyPage -= 1;
          loadLegacy();
        });
        pnext.addEventListener('click', function () {
          legacyPage += 1;
          loadLegacy();
        });

        box.appendChild(searchRow);
        box.appendChild(pager);
        box.appendChild(innerList);

        var opened = false;
        toggle.addEventListener('click', function () {
          opened = !opened;
          box.style.display = opened ? 'block' : 'none';
          toggle.textContent = opened ? '收起' : '展开';
          if (opened) loadLegacy();
        });

        li.appendChild(wrap);
        li.appendChild(toggle);
        li.appendChild(zipA);
        li.appendChild(box);
        signedListEl.appendChild(li);
      })();
    }

    if (batches && batches.length) {
      batches.forEach(function (b) {
        var li = document.createElement('li');
        var wrap = document.createElement('div');
        wrap.style.flex = '1';
        wrap.style.minWidth = '0';
        var title = document.createElement('div');
        title.style.fontWeight = '600';
        title.style.wordBreak = 'break-all';
        title.textContent =
          '批次 ' +
          String(b.batch_id || '').slice(0, 8) +
          (b.created_at ? (' · ' + b.created_at) : '') +
          '（' +
          (b.n || 0) +
          ' 个）';
        var meta = document.createElement('div');
        meta.className = 'signed-meta';
        meta.textContent = '可展开查看文件；可下载该批次 zip；批次内支持搜索。';
        wrap.appendChild(title);
        wrap.appendChild(meta);

        var toggle = document.createElement('button');
        toggle.type = 'button';
        toggle.className = 'btn btn-secondary';
        toggle.textContent = '展开';

        var zipA = document.createElement('a');
        zipA.className = 'btn btn-secondary';
        zipA.textContent = '下载批次包';
        zipA.href = apiUrl('/api/sign/signed-batch/' + b.batch_id + '/zip');
        zipA.setAttribute('download', '');

        var box = document.createElement('div');
        box.style.display = 'none';
        box.style.width = '100%';
        box.style.marginTop = '10px';
        box.style.paddingTop = '10px';
        box.style.borderTop = '1px solid var(--border)';

        var searchRow = document.createElement('div');
        searchRow.className = 'btn-row';
        searchRow.style.margin = '0 0 10px';
        var bi = document.createElement('input');
        bi.type = 'search';
        bi.placeholder = '在该批次内搜索文件名…';
        bi.style.flex = '1';
        bi.style.minWidth = '140px';
        bi.style.padding = '8px 12px';
        bi.style.border = '1px solid var(--border)';
        bi.style.borderRadius = '8px';
        var bs = document.createElement('button');
        bs.type = 'button';
        bs.className = 'btn btn-secondary';
        bs.textContent = '搜索';
        searchRow.appendChild(bi);
        searchRow.appendChild(bs);

        var innerList = document.createElement('ul');
        innerList.className = 'file-list';
        innerList.style.marginTop = '0';

        function loadBatchItems(q2) {
          innerList.innerHTML = '';
          var h = document.createElement('li');
          h.textContent = '正在加载批次文件…';
          h.style.border = 'none';
          h.style.background = 'transparent';
          innerList.appendChild(h);
          fetchJson(
            apiUrl(
              '/api/sign/signed-batch/' +
                b.batch_id +
                '?q=' +
                encodeURIComponent(q2 || '') +
                '&_=' +
                Date.now()
            )
          )
            .then(function (r) {
              var jj = r.data || {};
              if (!jj.ok) {
                innerList.innerHTML = '';
                var e = document.createElement('li');
                e.textContent = '加载失败：' + (jj.error || '未知错误');
                e.style.border = 'none';
                e.style.background = 'transparent';
                innerList.appendChild(e);
                return;
              }
              var arr = jj.items || [];
              innerList.innerHTML = '';
              if (!arr.length) {
                var e2 = document.createElement('li');
                e2.textContent = q2 ? '无匹配文件。' : '该批次暂无文件。';
                e2.style.border = 'none';
                e2.style.background = 'transparent';
                innerList.appendChild(e2);
                return;
              }
              arr.forEach(function (it) {
                renderSignedItemRow(it, innerList);
              });
            })
            .catch(function (e) {
              innerList.innerHTML = '';
              var e3 = document.createElement('li');
              e3.textContent = '加载失败：' + (e && e.message ? e.message : String(e));
              e3.style.border = 'none';
              e3.style.background = 'transparent';
              innerList.appendChild(e3);
            });
        }

        bs.addEventListener('click', function () {
          loadBatchItems(bi.value ? String(bi.value).trim() : '');
        });
        bi.addEventListener('keydown', function (ev) {
          if (ev && ev.key === 'Enter') {
            ev.preventDefault();
            loadBatchItems(bi.value ? String(bi.value).trim() : '');
          }
        });

        box.appendChild(searchRow);
        box.appendChild(innerList);

        var opened = false;
        toggle.addEventListener('click', function () {
          opened = !opened;
          box.style.display = opened ? 'block' : 'none';
          toggle.textContent = opened ? '收起' : '展开';
          if (opened) loadBatchItems('');
        });

        li.appendChild(wrap);
        li.appendChild(toggle);
        li.appendChild(zipA);
        li.appendChild(box);
        signedListEl.appendChild(li);
      });
      return;
    }

    items.forEach(function (it) {
      renderSignedItemRow(it, signedListEl);
    });
  }

  function refreshSignedList(opt) {
    opt = opt || {};
    if (!IS_FILE_SIGN_PAGE || !signedListEl || !signedHint) return Promise.resolve();
    if (!opt.skipPageProgress) beginPageProgress('正在加载已签名列表…');
    showSignedListLoading();
    var u =
      apiUrl('/api/sign/signed-batches') +
      '?q=' +
      encodeURIComponent(signedListQ) +
      '&page=' +
      signedListPage +
      '&page_size=' +
      signedListPageSize +
      '&_=' +
      Date.now();
    return fetchJson(u)
      .then(function (result) {
        var j = result.data;
        if (!j.ok) {
          signedListEl.innerHTML = '';
          signedHint.style.display = 'block';
          signedSearchRow.style.display = 'none';
          signedHint.textContent =
            '已签名列表加载失败：' + (j.error || '请稍后重试。');
          return;
        }
        renderSignedList(j);
      })
      .catch(function (e) {
        signedListEl.innerHTML = '';
        signedHint.style.display = 'block';
        signedSearchRow.style.display = 'none';
        signedHint.textContent =
          '已签名列表加载失败：' + (e && e.message ? e.message : String(e));
      })
      .finally(function () {
        if (!opt.skipPageProgress) endPageProgress();
      });
  }

  function renderStrokeItemList(j) {
    var items = (j && j.items) || [];
    var dbShare = !!(j && j.db_share);
    var total = typeof (j && j.total) === 'number' ? j.total : items.length;
    var page = typeof (j && j.page) === 'number' ? j.page : 1;
    var pageSize = typeof (j && j.page_size) === 'number' ? j.page_size : strokeItemPageSize;
    var strokePageCountEarly = Math.max(1, Math.ceil(total / pageSize) || 1);
    if (dbShare && total > 0 && page > strokePageCountEarly) {
      strokeItemPage = strokePageCountEarly;
      refreshStrokeItemList();
      return;
    }
    strokeItemListEl.innerHTML = '';
    if (!dbShare) {
      strokeItemsHint.style.display = 'block';
      strokeItemsHint.textContent =
        '当前未配置 MySQL 时，签字素材仅保存在会话目录，此列表不可用。配置 MYSQL_HOST 并保存笔迹后可在此检索与下载。';
      if (strokeItemPager) strokeItemPager.style.display = 'none';
      return;
    }
    if (strokeItemPager) strokeItemPager.style.display = '';
    if (!items.length) {
      strokeItemsHint.style.display = 'block';
      strokeItemsHint.textContent = strokeItemQ
        ? '无匹配的签字图片，请尝试其它关键字或清空搜索。'
        : '暂无已入库的签字图片。在上方保存笔迹后会出现于此。';
    } else {
      strokeItemsHint.style.display = 'none';
    }
    var pageCount = Math.max(1, Math.ceil(total / pageSize) || 1);
    strokeItemPagerInfo.textContent =
      '共 ' +
      total +
      ' 条 · 第 ' +
      page +
      ' / ' +
      pageCount +
      ' 页（每页 ' +
      pageSize +
      ' 条）';
    strokeItemPrevBtn.disabled = page <= 1;
    strokeItemNextBtn.disabled = page >= pageCount;
    items.forEach(function (it) {
      var li = document.createElement('li');
      var wrap = document.createElement('div');
      wrap.style.flex = '1';
      wrap.style.minWidth = '0';
      var t = document.createElement('div');
      t.style.fontWeight = '500';
      t.textContent =
        (it.signer_name || it.signer_id || '') +
        ' · ' +
        (it.kind_label || it.kind || '') +
        ' · ' +
        (it.locale === 'en' ? '英文' : '中文');
      var meta = document.createElement('div');
      meta.className = 'signed-meta';
      var mp = [];
      if (it.updated_at) mp.push(it.updated_at);
      if (it.id) mp.push('素材 id：' + it.id);
      var strokeFtp = ftpMetaLine(it.ftp_uploaded, it.ftp_last_error, '已保存在 MySQL，可下载');
      if (strokeFtp) mp.push(strokeFtp);
      meta.textContent = mp.join(' · ');
      wrap.appendChild(t);
      wrap.appendChild(meta);

      var dl = document.createElement('a');
      dl.className = 'btn btn-secondary';
      dl.href = apiUrl('/api/sign/stroke-items/' + it.id + '/png');
      dl.setAttribute('download', 'stroke-' + it.id + '.png');
      dl.textContent = '下载';

      var delBtn = document.createElement('button');
      delBtn.type = 'button';
      delBtn.className = 'btn btn-secondary del-btn';
      delBtn.textContent = '删除';
      delBtn.addEventListener('click', function () {
        if (!window.confirm('确定删除该条签字图片素材？已绑定到文件的映射会一并解除。')) return;
        withButtonBusy(delBtn, '删除中…', function () {
          return fetchJson(apiUrl('/api/sign/stroke-items/' + it.id), { method: 'DELETE' }).then(
            function (r) {
              var jj = r.data || {};
              if (!jj.ok) {
                setStrokeItemsActionFeedback(jj.error || '删除失败', true);
                return;
              }
              setStrokeItemsActionFeedback('', false);
              refreshStrokeItemList();
            }
          );
        }).catch(function (e) {
          setStrokeItemsActionFeedback(e.message || String(e), true);
        });
      });

      li.appendChild(wrap);
      li.appendChild(dl);
      li.appendChild(delBtn);
      strokeItemListEl.appendChild(li);
    });
  }

  function refreshStrokeItemList(opt) {
    opt = opt || {};
    if (!opt.skipPageProgress) beginPageProgress('正在加载签字素材列表…');
    showStrokeItemsLoading();
    var u =
      apiUrl('/api/sign/stroke-items') +
      '?q=' +
      encodeURIComponent(strokeItemQ) +
      (strokeItemCat ? '&cat=' + encodeURIComponent(strokeItemCat) : '') +
      '&page=' +
      strokeItemPage +
      '&page_size=' +
      strokeItemPageSize +
      '&_=' +
      Date.now();
    return fetchJson(u)
      .then(function (result) {
        var j = result.data;
        if (!j.ok) {
          strokeItemListEl.innerHTML = '';
          strokeItemsHint.style.display = 'block';
          if (strokeItemPager) strokeItemPager.style.display = 'none';
          strokeItemsHint.textContent =
            '签字素材列表加载失败：' + (j.error || '请稍后重试。');
          return;
        }
        renderStrokeItemList(j);
      })
      .catch(function (e) {
        strokeItemListEl.innerHTML = '';
        strokeItemsHint.style.display = 'block';
        if (strokeItemPager) strokeItemPager.style.display = 'none';
        strokeItemsHint.textContent =
          '签字素材列表加载失败：' + (e && e.message ? e.message : String(e));
      })
      .finally(function () {
        if (!opt.skipPageProgress) endPageProgress();
      });
  }

  function scheduleRefreshFileList(delayMs) {
    if (__refreshFileListDeferTimer) {
      clearTimeout(__refreshFileListDeferTimer);
    }
    __refreshFileListDeferTimer = setTimeout(function () {
      __refreshFileListDeferTimer = null;
      refreshFileList({ softFail: true }).catch(function () {});
    }, typeof delayMs === 'number' ? delayMs : 1500);
  }

  function refreshFileList(opt) {
    opt = opt || {};
    if (!IS_FILE_SIGN_PAGE || !fileListEl || !listHint) {
      return Promise.resolve();
    }
    if (__refreshFileListPromise && !opt.force) {
      return __refreshFileListPromise;
    }
    var myGen = ++_fileListRefreshGen;
    var prevFiles = (savedFiles || []).slice();
    if (!opt.skipPageProgress) {
      beginPageProgress(opt.progressLabel || '正在加载文件列表…');
    }
    if (!opt.silent) {
      showFileListLoading();
    }
    __refreshFileListPromise = fetchJsonWithRetry(
      apiUrl('/api/sign/files') + '?_=' + Date.now(),
      { cache: 'no-store', timeoutMs: SIGN_FILES_LIST_FETCH_TIMEOUT_MS },
      { maxTry: 2, delayMs: 1200 }
    )
      .then(function (result) {
        if (myGen !== _fileListRefreshGen) {
          return;
        }
        var j = result.data;
        if (!j.ok || !Array.isArray(j.files)) {
          if (opt.softFail && prevFiles.length) {
            savedFiles = prevFiles;
            listHint.style.display = 'block';
            listHint.textContent =
              '文件列表刷新失败' +
              (j && j.error ? '：' + j.error : '') +
              '，已保留当前页面上的文件记录。';
            renderFileList();
            return;
          }
          savedFiles = [];
          if (selectedFileId) selectedFileId = null;
          renderFileList();
          listHint.style.display = 'block';
          listHint.textContent =
            '文件列表加载失败' + (j && j.error ? ('：' + j.error) : '。请刷新页面或确认服务已启动。');
          return;
        }
        savedFiles = normalizeSavedFileRecords(j.files);
        if (
          selectedFileId &&
          !savedFiles.some(function (f) {
            return f.id === selectedFileId;
          })
        ) {
          selectedFileId = null;
        }
        listHint.style.display = 'none';
        renderFileList();
        if (isBatchWorkbenchMode()) {
          mergeBatchWorkbenchRowsFromSavedFiles();
          renderBatchWorkbenchTable();
        }
      })
      .catch(function (e) {
        if (myGen !== _fileListRefreshGen) {
          return;
        }
        if (opt.softFail && prevFiles.length) {
          savedFiles = prevFiles;
          listHint.style.display = 'block';
          listHint.textContent =
            '文件列表刷新超时（远程 MySQL 较慢）：' +
            (e && e.message ? e.message : String(e)) +
            '。已保留当前载入的文件，可稍后点「刷新列表」重试。';
          renderFileList();
          return;
        }
        savedFiles = [];
        if (selectedFileId) selectedFileId = null;
        renderFileList();
        listHint.style.display = 'block';
        listHint.textContent =
          '文件列表加载失败：' + (e && e.message ? e.message : String(e));
      })
      .finally(function () {
        if (!opt.skipPageProgress) endPageProgress();
        if (__refreshFileListPromise) {
          __refreshFileListPromise = null;
        }
      });
    return __refreshFileListPromise;
  }

  function setRoleChecked(roleId, checked) {
    var el = document.getElementById('chk_' + roleId);
    if (!el) return;
    el.checked = !!checked;
    var panel = document.getElementById('panel_' + roleId);
    if (panel) panel.classList.toggle('visible', !!checked);
  }

  /** 切换文件或清空列表时先取消所有签字角色勾选，避免沿用上一份文件的选中状态 */
  function resetAllRoleChecks() {
    ROLES.forEach(function (r) {
      setRoleChecked(r.id, false);
    });
  }

  function _decodeHandoffCtxB64(b64) {
    if (!b64 || typeof b64 !== 'string') return null;
    try {
      var raw = atob(b64.replace(/\s+/g, ''));
      var len = raw.length;
      var bytes = new Uint8Array(len);
      for (var i = 0; i < len; i++) {
        bytes[i] = raw.charCodeAt(i) & 0xff;
      }
      var txt;
      try {
        if (typeof TextDecoder !== 'undefined') {
          txt = new TextDecoder('utf-8').decode(bytes);
        } else {
          var bin = '';
          for (var k = 0; k < len; k++) {
            bin += String.fromCharCode(bytes[k]);
          }
          txt = decodeURIComponent(escape(bin));
        }
      } catch (_) {
        return null;
      }
      var o = JSON.parse(txt);
      return typeof o === 'object' && o ? o : null;
    } catch (_) {
      return null;
    }
  }

  function _normSignerHint(s) {
    var x = String(s || '')
      .replace(/（[^）]*）/g, '')
      .replace(/\([^)]*\)/g, '')
      .replace(/[\/／、，,;；|]+/g, ' ')
      .trim()
      .toLowerCase()
      .replace(/\s+/g, '');
    return x;
  }

  function _signerMatchTokens(hint) {
    var raw = String(hint || '').trim();
    if (!raw) return [];
    var parts = raw.split(/[\/／、，,;；|]+/).map(function (p) {
      return _normSignerHint(p);
    });
    var out = [];
    var seen = {};
    parts.forEach(function (p) {
      if (p && !seen[p]) {
        seen[p] = true;
        out.push(p);
      }
    });
    var whole = _normSignerHint(raw);
    if (whole && !seen[whole]) {
      out.unshift(whole);
    }
    return out;
  }

  function _findSignerById(sid) {
    var id = String(sid || '').trim();
    if (!id || !signersList || !signersList.length) return null;
    for (var i = 0; i < signersList.length; i++) {
      if (signersList[i] && String(signersList[i].id) === id) return signersList[i];
    }
    return null;
  }

  function _findSignerByNameHint(hint) {
    if (!hint || !signersList || !signersList.length) return null;
    var tokens = _signerMatchTokens(hint);
    if (!tokens.length) return null;
    var i;
    var k;
    for (k = 0; k < tokens.length; k++) {
      var t = tokens[k];
      if (!t) continue;
      for (i = 0; i < signersList.length; i++) {
        var s = signersList[i];
        var nm = _normSignerHint(s && s.name);
        if (nm && nm === t) return s;
      }
    }
    for (k = 0; k < tokens.length; k++) {
      t = tokens[k];
      if (!t) continue;
      for (i = 0; i < signersList.length; i++) {
        s = signersList[i];
        nm = _normSignerHint(s && s.name);
        if (nm && (nm.indexOf(t) >= 0 || t.indexOf(nm) >= 0)) return s;
      }
    }
    return null;
  }

  /** 将任务日期统一为 YYYY-MM-DD，供 date_iso 与正则使用 */
  function _parseAiwordDocDateIso(s) {
    var t = String(s || '').trim();
    if (!t) return '';
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(t);
    if (m) return t;
    m = /^(\d{4})[/.](\d{1,2})[/.](\d{1,2})$/.exec(t);
    if (m) {
      var y = m[1];
      var mo = ('0' + m[2]).slice(-2);
      var d = ('0' + m[3]).slice(-2);
      return y + '-' + mo + '-' + d;
    }
    m = /^(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日?$/.exec(t);
    if (m) {
      var y2 = m[1];
      var mo2 = ('0' + m[2]).slice(-2);
      var d2 = ('0' + m[3]).slice(-2);
      return y2 + '-' + mo2 + '-' + d2;
    }
    return '';
  }

  function _firstStrokeId(signer, kind, preferredLocale) {
    if (!signer) return '';
    var arr = kind === 'date' ? signer.date_items || [] : signer.sig_items || [];
    if (!arr.length || !arr[0]) return '';
    var pl =
      preferredLocale === 'en' || preferredLocale === 'zh' ? preferredLocale : null;
    if (pl) {
      for (var i = 0; i < arr.length; i++) {
        var it = arr[i] || {};
        if ((it.locale || 'zh') === pl) {
          return String(it.id || '');
        }
      }
    }
    return String(arr[0].id || '');
  }

  function _isLikelyEnglishCountry(s) {
    var raw = String(s || '').trim();
    if (!raw) return false; // 业务规则：未指明注册国家时，默认中文拼接
    // 含中文字符（如「中国 / 中華人民共和国 / 中国 NMPA」）=> 中文
    if (/[\u4e00-\u9fff]/.test(raw)) return false;
    var t = raw.toLowerCase();
    if (t.indexOf('nmpa') >= 0 || t.indexOf('china') >= 0) return false;
    return true;
  }

  /** aiword 入口：按项目注册国家默认「签字版本」（各角色 loc + 素材库 libLocale） */
  function applyAiwordHandoffSignLocaleDefaults(ctx) {
    if (!_isFromAiwordHandoff()) return;
    var c = ctx && typeof ctx === 'object' ? ctx : {};
    var loc = _isLikelyEnglishCountry(c.country != null ? c.country : '') ? 'en' : 'zh';
    try {
      ROLES.forEach(function (r) {
        if (!r || !r.id) return;
        roleLocaleMap[r.id] = loc;
      });
    } catch (_) {}
    try {
      if (libLocaleSelect) {
        libLocaleSelect.value = loc;
        if (typeof renderLibLocaleQuickPick === 'function') {
          renderLibLocaleQuickPick();
        }
      }
    } catch (_) {}
  }

  function clearAiwordHandoffState() {
    __aiwordHandoffCtx = null;
    __aiwordHandoffTargetFileId = null;
    __aiwordDeferDetectFileId = null;
    _hideAiwordHandoffLoadingMask();
  }

  function aiwordRolesFromCtx(ctx) {
    var ids = [];
    if (!ctx || typeof ctx !== 'object') return ids;
    var w = String(ctx.writer != null ? ctx.writer : '').trim();
    var e = String(ctx.editor != null ? ctx.editor : '').trim();
    if (w || e) {
      ids.push('author');
      ids.push('executor');
    }
    if (String(ctx.reviewer != null ? ctx.reviewer : '').trim()) ids.push('reviewer');
    if (String(ctx.approver != null ? ctx.approver : '').trim()) ids.push('approver');
    return ids;
  }

  function applyAiwordRoleChecksFromCtx(ctx) {
    // 按用户要求：签字位必须先由文档 detect 命中，再展示/自动勾选；
    // 因此 aiword 上下文不再直接驱动角色勾选。
    return;
  }

  function aiwordHandoffActiveRoleIds(ctx) {
    var active = [];
    var seen = {};
    function add(rid) {
      if (!rid || seen[rid]) return;
      seen[rid] = true;
      active.push(rid);
    }
    // 仅以 detect 命中角色为准，不再把 aiword 上下文角色并入 active。
    mergeDetectedRolesForUi().forEach(function (x) {
      add(x && x.id);
    });
    return active;
  }

  function registerAiwordHandoffFile(fid, ctxObj, filesOpt) {
    var rawCtx = ctxObj && typeof ctxObj === 'object' ? ctxObj : null;
    if (rawCtx && rawCtx.docDate != null && rawCtx.doc_date == null) {
      rawCtx = Object.assign({}, rawCtx, { doc_date: rawCtx.docDate });
    }
    var hasAnyHint =
      rawCtx &&
      (String(rawCtx.editor || '').trim() ||
        String(rawCtx.writer || '').trim() ||
        String(rawCtx.reviewer || '').trim() ||
        String(rawCtx.approver || '').trim() ||
        String(rawCtx.doc_date || '').trim() ||
        String(rawCtx.country || '').trim());
    selectedFileId = fid;
    if (Array.isArray(filesOpt) && filesOpt.length) {
      savedFiles = filesOpt;
    }
    __aiwordHandoffCtx = hasAnyHint ? rawCtx : null;
    __aiwordHandoffTargetFileId = hasAnyHint && fid ? fid : null;
    __aiwordDeferDetectFileId = fid || null;
    if (fid && rawCtx && typeof rawCtx === 'object') {
      __aiwordHandoffCtxByFileId[String(fid)] = rawCtx;
    }
    if (fid) {
      __aiwordHandoffAutoRedetectDoneFor[String(fid)] = false;
    }
    try {
      if (document.body && document.body.classList.contains('from-aiword')) {
        applyAiwordHandoffSignLocaleDefaults(rawCtx || {});
      }
    } catch (_) {}
  }

  function kickoffAiwordHandoffPipeline(fileId) {
    if (!fileId) return;
    if (!fileUiCache[fileId]) {
      fileUiCache[fileId] = {};
    }
    fileUiCache[fileId].detectedOnce = false;
    detectAndAutoSelectRoles(fileId, __aiwordHandoffDetectRetries);
  }

  function scheduleAiwordHandoffHintsAfterDetect(fileId) {
    if (!__aiwordHandoffCtx || fileId == null || __aiwordHandoffTargetFileId == null) {
      return Promise.resolve();
    }
    if (String(fileId) !== String(__aiwordHandoffTargetFileId)) {
      return Promise.resolve();
    }
    var ctxSnap = null;
    try {
      ctxSnap = JSON.parse(JSON.stringify(__aiwordHandoffCtx));
    } catch (_) {
      ctxSnap = __aiwordHandoffCtx;
    }
    if (!ctxSnap || typeof ctxSnap !== 'object') {
      return Promise.resolve();
    }
    return refreshSigners()
      .then(function () {
        if (String(selectedFileId) !== String(fileId)) return;
        return applyAiwordHandoffHintsOnce(fileId, ctxSnap).then(function () {
          setNeedSignActionFeedback('');
        });
      })
      .catch(function () {})
      .then(function () {
        if (!isBatchWorkbenchMode()) {
          clearAiwordHandoffState();
        }
      });
  }

  function applyAiwordHandoffHintsOnce(fileId, ctxOpt, opts) {
    opts = opts && typeof opts === 'object' ? opts : {};
    // 批量工作台：所有「按姓名匹配并写库」的工作都收敛到 workbenchSyncUiAfterMatch，
    // 它内部走 _workbenchPlanRoleMap 单次落库，避免与 detect/补日期/手动改版本互相覆盖。
    if (opts.batchSilent && isBatchWorkbenchMode()) {
      var signersReady = signersList && signersList.length;
      var chain = signersReady ? Promise.resolve() : refreshSigners();
      return chain.then(function () {
        return workbenchSyncUiAfterMatch(fileId);
      });
    }
    return refreshSigners().then(function () {
      var ctx =
        ctxOpt && typeof ctxOpt === 'object' ? ctxOpt : __aiwordHandoffCtx;
      if (!ctx || fileId == null) {
        return Promise.resolve();
      }
      if (!opts.skipSelectedCheck && String(selectedFileId) !== String(fileId)) {
        return Promise.resolve();
      }
      var writerHint = String(ctx.writer != null ? ctx.writer : '').trim();
      var editorHint = String(ctx.editor != null ? ctx.editor : '').trim();
      var compileHint = editorHint || writerHint;
      var compileHintAlt = writerHint && writerHint !== compileHint ? writerHint : '';
      var hintByRole = {
        author: compileHint,
        reviewer: String(ctx.reviewer != null ? ctx.reviewer : '').trim(),
        approver: String(ctx.approver != null ? ctx.approver : '').trim(),
        executor: compileHint,
      };
      var hintAltByRole = {
        author: compileHintAlt,
        executor: compileHintAlt,
      };
      var rowWb = opts.batchSilent ? __batchWorkbenchRows[String(fileId)] : null;
      var active = [];
      mergeDetectedRolesForFile(fileId).forEach(function (x) {
        if (x && x.id) active.push(x.id);
      });
      if (!active.length && !opts.batchSilent) {
        active = aiwordHandoffActiveRoleIds(ctx);
      }
      if (opts.batchSilent && rowWb) {
        ['author', 'reviewer', 'approver'].forEach(function (rid) {
          var h = _workbenchNameHintForRole(fileId, rowWb, rid) || hintByRole[rid];
          if (h && active.indexOf(rid) < 0) active.push(rid);
        });
      } else if (!active.length && opts.batchSilent) {
        ['author', 'reviewer', 'approver', 'executor'].forEach(function (rid) {
          if (hintByRole[rid] && active.indexOf(rid) < 0) active.push(rid);
        });
      }
      if (!active.length) {
        if (opts.batchSilent && isBatchWorkbenchMode()) {
          return workbenchSyncUiAfterMatch(fileId);
        }
        return Promise.resolve();
      }

      var docDate = _parseAiwordDocDateIso(ctx.doc_date != null ? ctx.doc_date : '');
      var isoOk = !!docDate;
      var useEnDate = rowWb
        ? rowWb.locale === 'en'
        : _isLikelyEnglishCountry(ctx.country);
      var preferredDateMode = useEnDate ? 'composite_en_space' : 'composite_zh_ymd';
      var wantStrokeLocale = useEnDate ? 'en' : 'zh';
      try {
        if (_isFromAiwordHandoff()) {
          applyAiwordHandoffSignLocaleDefaults(ctx);
        }
      } catch (_) {}

      var nextMap = opts.batchSilent
        ? _deepCloneJsonish(getFileRoleMapForWorkbench(fileId))
        : Object.assign({}, currentRoleMap || {});
      var touched = false;

      function _resolveSigner(rid) {
        var hint = hintByRole[rid] || '';
        if (!hint && opts.batchSilent && rowWb) {
          hint = _workbenchNameHintForRole(fileId, rowWb, rid);
        }
        var alt = hintAltByRole[rid] || '';
        var signer = hint ? _findSignerByNameHint(hint) : null;
        if (!signer && alt) signer = _findSignerByNameHint(alt);
        return { hint: hint || alt, signer: signer };
      }

      active.forEach(function (rid) {
        if (!rid) return;
        var resolved = _resolveSigner(rid);
        var hint = resolved.hint;
        var signer = resolved.signer;
        var p =
          nextMap[rid] && typeof nextMap[rid] === 'object' ? Object.assign({}, nextMap[rid]) : {};
        if (signer) {
          var strokeLoc =
            opts.batchSilent && rowWb
              ? _workbenchStrokeLocaleForRole(fileId, rid, rowWb)
              : wantStrokeLocale;
          _applyWorkbenchSignerMaterialToPair(p, signer, strokeLoc, isoOk ? docDate : '');
          if (opts.batchSilent && rowWb) {
            var fqF = _ensureRoleFilterStateForFile(fileId, rid);
            if (hint) fqF.sig = hint;
            fqF.date = _signerNameForFilter(signer, hint) || fqF.date || '';
          } else {
            var fq = _ensureRoleFilterState(rid);
            if (hint) fq.sig = hint;
            fq.date = _signerNameForFilter(signer, hint) || fq.date || '';
          }
        } else if (hint && roleMapEntryNonEmpty(p)) {
          /* 库中无此人但映射里已有部分素材：保留，不删除 */
        } else if (isoOk && !signersDbShare) {
          p.date_iso = docDate;
        }
        if (roleMapEntryNonEmpty(p)) {
          nextMap[rid] = p;
          touched = true;
        }
      });

      cachePatchCurrentRoleMap(fileId, nextMap);
      if (String(selectedFileId) === String(fileId)) {
        currentRoleMap = _deepCloneJsonish(nextMap);
      }

      if (!touched) {
        if (!opts.batchSilent) {
          renderNeedSignTable();
          updateSubmitState();
        } else if (isBatchWorkbenchMode()) {
          return workbenchSyncUiAfterMatch(fileId);
        }
        return Promise.resolve();
      }

      return fetchJson(apiUrl('/api/sign/files/' + fileId + '/role-map'), {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ map: nextMap }),
      })
        .then(function (r) {
          if (!opts.skipSelectedCheck && String(selectedFileId) !== String(fileId)) {
            if (opts.batchSilent && isBatchWorkbenchMode()) {
              return workbenchSyncUiAfterMatch(fileId);
            }
            return;
          }
          var jj = r.data || {};
          var finalMap = jj.ok ? jj.map || nextMap : nextMap;
          cachePatchCurrentRoleMap(fileId, finalMap);
          if (String(selectedFileId) === String(fileId)) {
            currentRoleMap = _deepCloneJsonish(finalMap);
          }
          if (!opts.batchSilent) {
            renderNeedSignTable();
            updateSubmitState();
          } else if (isBatchWorkbenchMode()) {
            return workbenchSyncUiAfterMatch(fileId);
          }
        })
        .catch(function (e) {
          if (opts.batchSilent && isBatchWorkbenchMode()) {
            setBatchWorkbenchMsg(
              '角色映射保存失败：' + ((e && e.message) || String(e || '')),
              true
            );
            return workbenchSyncUiAfterMatch(fileId);
          }
        });
    });
  }

  function detectAndAutoSelectRoles(fileId, retryLeft, opts) {
    opts = opts && typeof opts === 'object' ? opts : {};
    retryLeft = typeof retryLeft === 'number' ? retryLeft : 0;
    if (!fileId) return opts.returnPromise ? Promise.resolve() : false;
    var detectDoneResolve = null;
    var detectDonePromise = null;
    if (opts.returnPromise) {
      detectDonePromise = new Promise(function (resolve) {
        detectDoneResolve = resolve;
      });
      fileUiCache[fileId] = fileUiCache[fileId] || {};
      fileUiCache[fileId]._detectInflightPromise = detectDonePromise;
      detectDonePromise.finally(function () {
        try {
          if (fileUiCache[fileId]) {
            delete fileUiCache[fileId]._detectInflightPromise;
          }
        } catch (_) {}
      });
    }
    function finishDetect(ok) {
      if (detectDoneResolve) {
        detectDoneResolve(ok !== false);
        detectDoneResolve = null;
      }
    }
    if (String(detectInFlightFor) === String(fileId)) {
      var inflight = (fileUiCache[fileId] || {})._detectInflightPromise;
      if (opts.returnPromise && inflight) {
        return inflight;
      }
      return opts.returnPromise ? detectDonePromise : false;
    }
    detectInFlightFor = fileId;
    var myEpoch = ++detectEpoch;
    var seq = ++detectRequestSeq;
    var myFileDetectToken = ++fileDetectRequestToken;
    fileUiCache[fileId] = fileUiCache[fileId] || {};
    fileUiCache[fileId]._detectToken = myFileDetectToken;
    lastDetectData = null;
    lastDetectFileId = null;
    lastDetectError = '';
    resetAllRoleChecks();
    var trackDetectProgress = !opts.skipPageProgress && !opts.batchSilent;
    var detectRetryScheduled = false;
    if (trackDetectProgress) {
      beginPageProgress(
        '正在识别需签角色与签字位：' + (_fileDisplayNameById(fileId) || fileId)
      );
    }
    if (redetectRolesBtn) {
      redetectRolesBtn.disabled = true;
      redetectRolesBtn.innerHTML =
        '<span class="spinner" aria-hidden="true"></span> 分析中…';
    }
    if (!opts.batchSilent) {
      needSignTable.innerHTML = '';
      needSignTable.textContent =
        retryLeft < __aiwordHandoffDetectRetries
          ? '正在分析模板与角色映射（重试 ' +
            (__aiwordHandoffDetectRetries - retryLeft) +
            '/' +
            __aiwordHandoffDetectRetries +
            '）…'
          : '正在分析模板与角色映射…';
    }
    fetchJson(apiUrl('/api/sign/detect?file_id=' + encodeURIComponent(fileId)), {
      timeoutMs: opts.batchSilent
        ? Math.min(_detectTimeoutMs(), _batchDetectPerFileTimeoutMs() + 90000)
        : _detectTimeoutMs(),
    })
      .then(function (result) {
        var stTok = (fileUiCache[fileId] || {})._detectToken;
        if (stTok !== myFileDetectToken) {
          finishDetect(false);
          return { __abort: true };
        }
        if (!opts.batchSilent && String(selectedFileId) !== String(fileId)) {
          finishDetect(false);
          return { __abort: true };
        }
        var j = result.data || {};
        var detectMismatch = j.ok ? validateDetectResponseForFile(fileId, j) : '';
        if (j.ok && !detectMismatch) {
          lastDetectError = '';
          lastDetectData = j;
          lastDetectFileId = fileId;
          cacheDetectResultForFile(fileId, j, '');
        } else if (j.ok && detectMismatch) {
          lastDetectData = null;
          lastDetectFileId = null;
          lastDetectError = detectMismatch;
          cacheDetectResultForFile(fileId, null, detectMismatch);
        } else {
          lastDetectData = null;
          lastDetectFileId = null;
          lastDetectError = (j && j.error) || '识别接口返回失败';
          cacheDetectResultForFile(fileId, null, lastDetectError);
          if (
            retryLeft > 0 &&
            __aiwordHandoffTargetFileId != null &&
            String(fileId) === String(__aiwordHandoffTargetFileId)
          ) {
            var delayMs = 450 + (__aiwordHandoffDetectRetries - retryLeft) * 400;
            detectRetryScheduled = true;
            setTimeout(function () {
              if (String(selectedFileId) !== String(fileId)) return;
              if (myEpoch !== detectEpoch) return;
              detectInFlightFor = null;
              detectAndAutoSelectRoles(fileId, retryLeft - 1, opts);
            }, delayMs);
            return { __retrying: true, __abort: true };
          }
        }
        // 这里曾用 String(...) 包左边，导致字符串与数字 !== 永远成立，detect 结果被
        // 静默丢弃 → 用户经常看到「第一次识别错、第二次重识别才对」。
        if ((fileUiCache[fileId] || {})._detectToken !== myFileDetectToken) {
          finishDetect(false);
          return { __abort: true };
        }
        cacheMarkDetected(fileId);
        var roles = [];
        if (j.ok) {
          roles = mergeDetectedRolesForFile(fileId);
          if (!roles.length && !opts.batchSilent) roles = mergeDetectedRolesForUi();
          if (roles.length) {
            roles.forEach(function (r) {
              if (r && r.id) setRoleChecked(r.id, true);
            });
          }
        }
        var resizeIds = roles
          .map(function (x) {
            return x && x.id;
          })
          .filter(Boolean);
        if (resizeIds.length) {
          requestAnimationFrame(function () {
            requestAnimationFrame(function () {
              resizeCanvasesForRoles(resizeIds);
            });
          });
        }
        // batchSilent：processBatchWorkbenchFile 已经在 ensureFileRoleMapLoaded 阶段
        // 拉过 role-map 写入 fileUiCache，这里不再重复发一次请求（节省每文件一次 HTTP 往返）。
        if (opts.batchSilent) {
          return { __abort: false };
        }
        return fetchJson(apiUrl('/api/sign/files/' + fileId + '/role-map')).then(function (rm) {
          if (String(selectedFileId) !== String(fileId)) {
            return { __abort: true };
          }
          if (rm && rm.data && rm.data.ok) {
            var rmMap = rm.data.map || {};
            if (String(selectedFileId) === String(fileId)) {
              currentRoleMap = rmMap;
            }
            cachePatchCurrentRoleMap(fileId, rmMap);
          }
          return { __abort: false };
        });
      })
      .then(function (pack) {
        if (pack && (pack.__abort || pack.__retrying)) {
          if (pack.__abort) finishDetect(false);
          return;
        }
        var stDet = (fileUiCache[fileId] || {}).lastDetectData;
        var detectOk = !!(stDet && stDet.ok);
        if (detectOk && !mergeDetectedRolesForFile(fileId).length) {
          detectOk = false;
          if (!fileUiCache[fileId].lastDetectError) {
            fileUiCache[fileId].lastDetectError =
              '未在文档中匹配到编制/编写/审核/批准或 Author/Reviewer/Approver 等签字标签';
          }
        }
        if (detectOk) {
          var probe = stDet && stDet.slot_probe && typeof stDet.slot_probe === 'object'
            ? stDet.slot_probe
            : null;
          if (probe && probe.ok === false) {
            detectOk = false;
            var miss = Array.isArray(probe.missing_roles)
              ? probe.missing_roles.map(function (x) { return roleLabel(x); }).join('、')
              : '';
            fileUiCache[fileId].lastDetectError =
              '已识别到角色，但签字位可落位校验未通过' +
              (miss ? '（未落位：' + miss + '）' : '') +
              (probe.error ? '：' + String(probe.error) : '');
            cacheDetectResultForFile(fileId, stDet, fileUiCache[fileId].lastDetectError);
          }
        }
        if (!opts.batchSilent) {
          renderNeedSignTable();
          updateSubmitState();
          if (!detectOk) {
            finishDetect(false);
            return;
          }
          return scheduleAiwordHandoffHintsAfterDetect(fileId).then(function () {
            finishDetect(true);
          });
        }
        saveCurrentFileUiToCache(fileId);
        refreshWorkbenchRowMaterial(fileId);
        finishDetect(detectOk);
      })
      .catch(function (err) {
        if (!opts.batchSilent && String(selectedFileId) !== String(fileId)) return;
        // 已超时或被取消的请求，重试一次仍然会超时/取消（耗时来自后端处理本身），
        // 直接走「识别失败」终态，让批量 pipeline 继续往下，避免反复刷请求。
        var msg = String((err && err.message) || err || '');
        var isAbortLike = /超时|timeout|已取消上传|aborted/i.test(msg);
        if (
          !isAbortLike &&
          retryLeft > 0 &&
          __aiwordHandoffTargetFileId != null &&
          String(fileId) === String(__aiwordHandoffTargetFileId)
        ) {
          detectRetryScheduled = true;
          var delayMs2 = 450 + (__aiwordHandoffDetectRetries - retryLeft) * 400;
          setTimeout(function () {
            if (String(selectedFileId) !== String(fileId)) return;
            if (myEpoch !== detectEpoch) return;
            detectInFlightFor = null;
            detectAndAutoSelectRoles(fileId, retryLeft - 1, opts);
          }, delayMs2);
          return;
        }
        lastDetectData = null;
        lastDetectFileId = null;
        lastDetectError =
          (err && err.message) ||
          '识别请求失败（网络或服务异常）。请稍后重试「重新识别」。';
        cacheDetectResultForFile(fileId, null, lastDetectError);
        if (!opts.batchSilent) {
          renderNeedSignTable();
          updateSubmitState();
          return scheduleAiwordHandoffHintsAfterDetect(fileId).then(function () {
            finishDetect(false);
          });
        }
        finishDetect(false);
      })
      .then(function () {
        if (myEpoch === detectEpoch) {
          detectInFlightFor = null;
        }
        if (seq === detectRequestSeq && redetectRolesBtn) {
          redetectRolesBtn.disabled = false;
          redetectRolesBtn.innerHTML = redetectRolesBtnDefaultHtml;
        }
      })
      .finally(function () {
        if (trackDetectProgress && !detectRetryScheduled) endPageProgress();
      });
    if (opts.returnPromise) return detectDonePromise;
    return true;
  }

  function manualRedetectNeedSignRoles() {
    setNeedSignActionFeedback('');
    if (!selectedFileId) {
      setNeedSignActionFeedback('请先在上方的文件列表中选择一项。');
      return;
    }
    if (String(detectInFlightFor) === String(selectedFileId)) {
      setNeedSignActionFeedback('正在分析当前文件，请稍候再试。');
      return;
    }
    lastDetectData = null;
    lastDetectFileId = null;
    lastDetectError = '';
    detectAndAutoSelectRoles(selectedFileId);
  }

  if (redetectRolesBtn) {
    redetectRolesBtn.addEventListener('click', function () {
      withButtonBusy(redetectRolesBtn, '识别中…', function () {
        manualRedetectNeedSignRoles();
        return Promise.resolve();
      }, { skipPageProgress: true });
    });
  }

  function ensureBatchApplyRoleMapBtn() {
    if (batchApplyRoleMapBtn) return batchApplyRoleMapBtn;
    if (!IS_FILE_SIGN_PAGE) return null;
    if (!redetectRolesBtn || !redetectRolesBtn.parentNode) return null;
    // 插到“重新识别”按钮左侧（同一按钮组）
    var b = document.createElement('button');
    b.type = 'button';
    b.className = 'btn btn-secondary';
    b.id = 'batchApplyRoleMapBtn';
    b.textContent = '将本文件映射批量应用到已勾选文件';
    try {
      redetectRolesBtn.parentNode.insertBefore(b, redetectRolesBtn);
    } catch (_) {
      redetectRolesBtn.parentNode.appendChild(b);
    }
    batchApplyRoleMapBtn = b;
    return b;
  }

  function ensureSaveRoleConfigBtn() {
    if (saveRoleConfigBtn) return saveRoleConfigBtn;
    if (!IS_FILE_SIGN_PAGE) return null;
    if (!redetectRolesBtn || !redetectRolesBtn.parentNode) return null;
    var b = document.createElement('button');
    b.type = 'button';
    b.className = 'btn btn-secondary';
    b.id = 'saveRoleConfigBtn';
    b.textContent = '保存角色与签署人配置';
    b.title =
      '将当前文件「需签角色」勾选与「角色→素材」表完整写入服务器。库映射生成前也会自动保存；改配置后点一次可确认已生效。';
    try {
      redetectRolesBtn.parentNode.insertBefore(b, redetectRolesBtn);
    } catch (_) {
      redetectRolesBtn.parentNode.appendChild(b);
    }
    saveRoleConfigBtn = b;
    return b;
  }

  function _checkedBatchFileIds() {
    var root = null;
    if (isBatchWorkbenchMode() && batchWorkbenchHiddenPicks) {
      root = batchWorkbenchHiddenPicks;
    }
    var nodes = root
      ? root.querySelectorAll('.batch-pick:checked')
      : document.querySelectorAll('.batch-pick:checked');
    var seen = {};
    var out = [];
    Array.from(nodes).forEach(function (el) {
      var id = el.getAttribute('data-id');
      if (!id || seen[String(id)]) return;
      seen[String(id)] = true;
      out.push(String(id));
    });
    return out;
  }

  function batchApplyCurrentRoleMapToPickedFiles() {
    setNeedSignActionFeedback('');
    if (!selectedFileId) {
      setNeedSignActionFeedback('请先选择一个文件（作为映射来源）。');
      return;
    }
    var ids = _checkedBatchFileIds();
    if (!ids.length) {
      setNeedSignActionFeedback('请先在文件列表勾选要批量应用的文件。');
      return;
    }
    // 不包含自己也没关系；包含则相当于重新保存一次
    if (!window.confirm('将“当前文件”的角色→素材映射，批量覆盖到已勾选的 ' + ids.length + ' 个文件？')) {
      return;
    }
    var mPreview = _deepCloneJsonish(currentRoleMap || {});
    if (!Object.keys(mPreview).length) {
      setNeedSignActionFeedback('当前文件尚未绑定任何映射（表格里未选择签名/日期素材）。');
      return;
    }
    var srcRec = savedFiles.find(function (x) {
      return x && String(x.id) === String(selectedFileId);
    });
    var srcName = (srcRec && (srcRec.name || srcRec.id)) ? String(srcRec.name || srcRec.id) : String(selectedFileId);
    if (batchApplyRoleMapBtn) {
      batchApplyRoleMapBtn.disabled = true;
      batchApplyRoleMapBtn.innerHTML = '<span class="spinner" aria-hidden="true"></span> 应用中…';
    }
    persistRoleMapToServer(selectedFileId)
      .then(function () {
        var m0 = _deepCloneJsonish(currentRoleMap || {});
        if (!Object.keys(m0).length) {
          setNeedSignActionFeedback('当前文件尚未绑定任何映射（表格里未选择签名/日期素材）。');
          return Promise.reject(new Error('empty_map'));
        }
        // 逐个 PUT（避免后端新增接口；同时可在失败时精确提示文件）
        var okN = 0;
        var fail = [];
        var chain = Promise.resolve();
        ids.forEach(function (fid) {
          chain = chain.then(function () {
            return fetchJson(apiUrl('/api/sign/files/' + fid + '/role-map'), {
              method: 'PUT',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ map: m0 }),
            }).then(function (r) {
              var jj = r.data || {};
              if (jj.ok) {
                okN += 1;
                cachePatchCurrentRoleMap(fid, jj.map || m0);
              } else {
                fail.push((fid || '') + '：' + (jj.error || '失败'));
              }
            }).catch(function (e) {
              fail.push((fid || '') + '：' + (e && e.message ? e.message : String(e)));
            });
          });
        });
        return chain.then(function () {
          var okMsg = '已将“' + srcName + '”文件角色配置批量映射到 ' + okN + ' 个文件。';
          ids.forEach(function (fid2) {
            if (!fid2) return;
            var isFailed = fail.some(function (x) {
              return String(x || '').indexOf(String(fid2) + '：') === 0;
            });
            if (!isFailed) cacheSetNeedSignNotice(fid2, okMsg, false);
          });
          if (ids.indexOf(String(selectedFileId)) < 0) {
            cacheSetNeedSignNotice(selectedFileId, '', false);
          }
          if (fail.length) {
            setNeedSignActionFeedback(
              '批量映射完成：成功 ' +
                okN +
                ' 个；失败 ' +
                fail.length +
                ' 个：' +
                fail.slice(0, 3).join('；') +
                (fail.length > 3 ? '…' : ''),
              true
            );
          } else {
            if (ids.indexOf(String(selectedFileId)) >= 0) {
              setNeedSignActionFeedback(okMsg, false);
            } else {
              setNeedSignActionFeedback('');
            }
          }
          renderNeedSignTable();
          updateSubmitState();
        });
      })
      .catch(function (e) {
        if (!e || e.message !== 'empty_map') {
          setNeedSignActionFeedback(e && e.message ? e.message : String(e), true);
        }
      })
      .then(function () {
        if (batchApplyRoleMapBtn) {
          batchApplyRoleMapBtn.disabled = false;
          batchApplyRoleMapBtn.textContent = '将本文件映射批量应用到已勾选文件';
        }
      });
  }

  if (batchApplyRoleMapBtn) {
    batchApplyRoleMapBtn.addEventListener('click', function () {
      batchApplyCurrentRoleMapToPickedFiles();
    });
  }

  // 兼容旧版 sign.html：若缺少按钮节点，自动插入并启用
  ensureBatchApplyRoleMapBtn();
  if (batchApplyRoleMapBtn && !batchApplyRoleMapBtn.__boundApplyRoleMap) {
    batchApplyRoleMapBtn.__boundApplyRoleMap = true;
    batchApplyRoleMapBtn.addEventListener('click', function () {
      batchApplyCurrentRoleMapToPickedFiles();
    });
  }
  ensureSaveRoleConfigBtn();
  if (saveRoleConfigBtn && !saveRoleConfigBtn.__boundSaveRoleCfg) {
    saveRoleConfigBtn.__boundSaveRoleCfg = true;
    saveRoleConfigBtn.addEventListener('click', function () {
      if (!selectedFileId) {
        setNeedSignActionFeedback('请先在文件列表中选择一项。', true);
        return;
      }
      setNeedSignActionFeedback('');
      withButtonBusy(saveRoleConfigBtn, '保存中…', function () {
        return persistRoleMapToServer(selectedFileId)
          .then(function () {
            saveCurrentFileUiToCache(selectedFileId);
            setNeedSignActionFeedback('已保存本文件的角色与签署人映射（含表格中当前选择）。', false);
          })
          .catch(function (e) {
            setNeedSignActionFeedback(e.message || String(e), true);
          });
      });
    });
  }

  /** 「生成已签名文档 / 批量签」错误（#errMsg，在「签名来源」区块下方）；其它操作请用 setNeedSignActionFeedback / setRoleRowFeedback 等 */
  function showErr(s) {
    if (s) {
      if (signerErrMsg) {
        signerErrMsg.style.display = 'none';
        signerErrMsg.textContent = '';
      }
      clearNeedSignScopedFeedbacks();
    }
    if (!errMsg) return;
    errMsg.style.display = s ? 'block' : 'none';
    errMsg.textContent = s || '';
  }

  function isCanvasBlank(canvas) {
    if (!canvas) return true;
    var w = canvas.width;
    var h = canvas.height;
    if (!w || !h) return true;
    var ctx = canvas.getContext('2d');
    var data;
    try {
      data = ctx.getImageData(0, 0, w, h).data;
    } catch (_) {
      return false;
    }
    // 移动端高 DPI + 抗锯齿/阴影边缘常见极低 alpha，阈值过严会误判「未签名」
    for (var i = 0; i < data.length; i += 4) {
      var r = data[i];
      var g = data[i + 1];
      var b = data[i + 2];
      var a = data[i + 3];
      if (a > 2) return false;
      if (a > 0 && r + g + b < 748) return false;
    }
    return true;
  }

  if (fileInput) {
    fileInput.addEventListener('change', function () {
      mergePendingSignFiles(filterSignFiles(Array.from(fileInput.files || [])));
      fileInput.value = '';
      updatePendingHint();
    });
  }

  if (dirInput) {
    dirInput.addEventListener('change', function () {
      mergePendingSignFiles(filterSignFiles(Array.from(dirInput.files || [])));
      dirInput.value = '';
      updatePendingHint();
    });
  }

  if (saveBtn) saveBtn.addEventListener('click', function () {
    if (__signUploadInFlight) return;
    clearFileRegionErr();
    setSaveUploadFeedback('');
    if (!pendingSignFiles.length) return;
    __signUploadInFlight = true;
    var form = new FormData();
    var nPending = pendingSignFiles.length;
    var hasArchivePending = pendingSelectionHasArchive();
    pendingSignFiles.forEach(function (f) {
      var name =
        f.webkitRelativePath && String(f.webkitRelativePath).length
          ? f.webkitRelativePath
          : f.name;
      form.append('files', f, name);
    });
    var uploadTimeoutMs = _saveUploadTimeoutMs();
    var progressTicker = null;
    var uploadStarted = Date.now();
    function tickUploadProgress() {
      var elapsed = Math.round((Date.now() - uploadStarted) / 1000);
      var phase = hasArchivePending
        ? '正在上传并解压压缩包（' + nPending + ' 项，已等待 ' + elapsed + ' 秒）…'
        : '正在上传并保存到列表（' + nPending + ' 个文件，已等待 ' + elapsed + ' 秒）…';
      var pct = Math.min(92, 8 + Math.floor(elapsed / 3));
      setSaveUploadProgress(true, pct, phase);
      if (fileHint) fileHint.textContent = phase;
    }
    beginPageProgress(hasArchivePending ? '上传并解压压缩包…' : '上传并保存到列表…');
    updatePageProgress(hasArchivePending ? '上传并解压压缩包…' : '上传并保存到列表…');
    progressTicker = setInterval(function () {
      tickUploadProgress();
      updatePageProgress(
        (hasArchivePending ? '上传并解压压缩包' : '上传并保存到列表') +
          '（' +
          Math.round((Date.now() - uploadStarted) / 1000) +
          ' 秒）'
      );
    }, 1200);
    tickUploadProgress();
    return withButtonBusy(saveBtn, hasArchivePending ? '上传解压中…' : '上传中…', function () {
      return fetchJson(apiUrl('/api/sign/upload'), {
        method: 'POST',
        body: form,
        timeoutMs: uploadTimeoutMs,
      }).then(
        function (result) {
          updatePageProgress('上传完成，正在同步文件列表…');
          var j = result.data;
          if (!j.ok) {
            setSaveUploadFeedback(j.error || '保存失败', 'error');
            throw new Error(j.error || '保存失败');
          }
          var warns = Array.isArray(j.warnings) ? j.warnings.filter(Boolean) : [];
          var arcLine = formatArchiveSummaryLine(j.archive_summary);
          if (warns.length) {
            var fb = warns.slice(0, 2).join('；');
            if (arcLine) fb = arcLine + '；' + fb;
            setSaveUploadFeedback(fb, 'warn');
          } else if (arcLine) {
            setSaveUploadFeedback(arcLine, 'ok');
          } else {
            setSaveUploadFeedback('已保存 ' + (j.added_ids && j.added_ids.length ? j.added_ids.length : nPending) + ' 个文件到列表', 'ok');
          }
          savedFiles = j.files || [];
          var addedIds = Array.isArray(j.added_ids) ? j.added_ids.map(String).filter(Boolean) : [];
          var hasArchive = !!j.upload_has_archive || (Number(j.archive_expanded) || 0) > 0;
          selectedFileId =
            (j.file && j.file.id) ||
            (savedFiles.length && savedFiles[savedFiles.length - 1].id);
          pendingSignFiles = [];
          fileInput.value = '';
          if (dirInput) dirInput.value = '';
          fileHint.textContent = '已保存，可继续添加；正在打开批量工作台并刷新列表…';
          saveBtn.disabled = true;
          enableBatchWorkbenchMode();
          var wbIds = addedIds.length
            ? addedIds
            : (selectedFileId ? [String(selectedFileId)] : []);
          return refreshWorkbenchFilesFromServer({ silent: false, skipPageProgress: true }).then(function () {
            syncHiddenBatchPicks();
            renderBatchWorkbenchTable();
            var needDetect = wbIds.some(function (id) {
              return !fileHasValidDetectCache(id);
            });
            if (wbIds.length && needDetect && !__batchWorkbenchPipelineRunning) {
              return startManualUploadBatchWorkbench(wbIds, {
                skipFileListRefresh: true,
                skipDetectIfCached: true,
              });
            }
            if (hasArchive && !wbIds.length) {
              setBatchWorkbenchMsg('压缩包已处理，但未得到可签字文档。请检查格式或查看上传提示。', 'warn');
            }
          });
        }
      );
    }, { skipRestoreDisabled: true, skipPageProgress: true })
      .catch(function (e) {
        setSaveUploadFeedback(e.message || String(e), 'error');
      })
      .finally(function () {
        __signUploadInFlight = false;
        if (progressTicker) clearInterval(progressTicker);
        setSaveUploadProgress(false);
        endPageProgress();
        saveBtn.disabled = !pendingSignFiles.length;
        if (!pendingSignFiles.length && fileHint) {
          fileHint.textContent =
            '已保存，可继续添加；批量工作台已更新，已识别记录将自动恢复';
        }
      });
  });

  if (submitBtn) submitBtn.addEventListener('click', function () {
    showErr('');
    var batchMode = !!(batchModeCb && batchModeCb.checked);
    var wbMode = isBatchWorkbenchMode();
    var pickedIds = _checkedBatchFileIds();
    // 工作台勾选了多个文件时，底部按钮与「一键批量签字」一致，按勾选列表批量生成
    if (wbMode && pickedIds.length > 0) {
      syncHiddenBatchPicks();
      if (batchModeCb) batchModeCb.checked = true;
      if (signSourceMode) signSourceMode.value = 'library';
      saveCurrentFileCanvasToCache(selectedFileId);
      _doBatchSignFromSubmit({
        apply_person: true,
        apply_date: true,
        workbenchMode: true,
        feedbackTarget: 'workbench',
      });
      return;
    }
    if (batchMode) {
      _doBatchSignFromSubmit();
      return;
    }
    if (!selectedFileId) {
      showErr('请先在列表中选择要签名的文件');
      return;
    }
    var rec = savedFiles.find(function (x) {
      return x.id === selectedFileId;
    });
    if (!rec) {
      showErr('未找到所选文件，请刷新页面后重试');
      return;
    }
    var source2 = signSourceValue();
    var roles = effectiveRolesForSingleSign();
    if (!roles.length) {
      showErr(
        source2 === 'library' && !libraryRolesRestrictedToChecks()
          ? '请在本文件需签角色表中至少绑定一项可生成的签名或日期；或在上方「签字角色（可多选）」卡片勾选要套打的角色，并勾选本处「与上方该卡片勾选一致」高级选项。'
          : '请至少勾选一个角色'
      );
      return;
    }
    if (source2 === 'canvas') {
      resizeCanvasesForRoles(roles);
    }

    if (source2 === 'library' && signersDbShare) {
      var preOne = buildSignPreflightForFiles([selectedFileId]);
      if (!preOne.canProceed) {
        showErr('当前文件没有任何角色已绑定可用素材，请先匹配签名/日期或用手写补签');
        return;
      }
      if (
        (preOne.partial.length || preOne.compositeWarn.length) &&
        !confirmPartialSignProceed(preOne)
      ) {
        setNeedSignActionFeedback('已取消生成');
        return;
      }
    }

    var persistPre = source2 === 'library' ? persistRoleMapToServer(selectedFileId) : Promise.resolve();
    beginPageProgress('正在生成已签名文档…');
    submitBtn.disabled = true;
    submitBtn.innerHTML =
      source2 === 'library'
        ? '<span class="spinner"></span> 保存配置并生成…'
        : '<span class="spinner"></span> 处理中…';

    persistPre
      .then(function () {
        var form = new FormData();
        form.append('file_id', selectedFileId);
        form.append('roles', JSON.stringify(roles));
        form.append('sign_source', source2);
        if (source2 === 'canvas') {
          roles.forEach(function (id) {
            var sigC = document.getElementById('sig_' + id);
            var dateC = document.getElementById('date_' + id);
            if (sigC && !isCanvasBlank(sigC)) {
              form.append('sig_' + id, _normalizedPngDataUrl(sigC, 'sig'));
            }
            if (dateC && !isCanvasBlank(dateC)) {
              form.append('date_' + id, _normalizedPngDataUrl(dateC, 'date'));
            }
          });
        }
        return fetch(apiUrl('/api/sign'), { method: 'POST', body: form, credentials: 'include' })
      .then(function (res) {
        var ct = res.headers.get('Content-Type') || '';
        if (!res.ok) {
          return res.text().then(function (errText) {
            if (ct.indexOf('application/json') !== -1) {
              try {
                var j = JSON.parse(errText);
                showErr(j.error || res.statusText);
              } catch (_) {
                showErr(
                  errText.trim().charAt(0) === '<'
                    ? '请求失败（HTTP ' + res.status + '），请确认服务已重启'
                    : errText.slice(0, 200)
                );
              }
            } else {
              showErr(
                errText.trim().charAt(0) === '<'
                  ? '请求失败（HTTP ' + res.status + '）'
                  : errText.slice(0, 200) || res.statusText
              );
            }
          });
        }
        if (ct.indexOf('application/json') !== -1) {
          return res.text().then(function (errText) {
            try {
              var j2 = JSON.parse(errText);
              showErr(j2.error || '失败');
            } catch (_) {
              showErr('失败');
            }
          });
        }
        return res.blob().then(function (blob) {
          try {
            var sumB64 = res.headers.get('X-Sign-Apply-Summary-B64') || '';
            if (sumB64) {
              var sumJson = '';
              sumJson = decodeURIComponent(
                Array.prototype.map
                  .call(atob(sumB64), function (c) {
                    return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
                  })
                  .join('')
              );
              var sum = JSON.parse(sumJson) || {};
              showBatchResult(formatSingleSignApplySummary(sum), false);
            } else {
              showBatchResult(
                '已生成文档。若部分角色未选择素材或拼接笔迹未录入，将自动跳过对应字段。',
                false
              );
            }
          } catch (_) {
            showBatchResult(
              '已生成文档。若部分角色未选择素材或拼接笔迹未录入，将自动跳过对应字段。',
              false
            );
          }
          var dispo = res.headers.get('Content-Disposition') || '';
          var stem = (rec.name || 'document').replace(/\.[^.]+$/, '');
          var extPart = '.docx';
          if (rec.name && /\.[^.]+$/.test(rec.name)) {
            extPart = (rec.name.match(/\.[^.]+$/) || ['.docx'])[0];
          } else if (rec.ext) {
            extPart = rec.ext.charAt(0) === '.' ? rec.ext : '.' + rec.ext;
          }
          var name = stem + '_signed' + extPart;
          var m = /filename\*?=(?:UTF-8'')?["']?([^"';]+)/i.exec(dispo);
          if (m) {
            try {
              name = decodeURIComponent(m[1].replace(/['"]/g, ''));
            } catch (_) {}
          }
          var a = document.createElement('a');
          a.href = URL.createObjectURL(blob);
          a.download = name;
          a.click();
          URL.revokeObjectURL(a.href);
          refreshSignedList({ skipPageProgress: true });
        });
      });
    })
      .catch(function (e) {
        showErr(e.message || String(e));
      })
      .then(function () {
        submitBtn.disabled = false;
        submitBtn.textContent = '生成已签名文档';
        updateSubmitState();
      })
      .finally(function () {
        endPageProgress();
      });
  });

  function _handoffClaimFetch(url) {
    return fetchJsonWithRetry(
      url,
      {
        method: 'POST',
        credentials: 'same-origin',
        timeoutMs: HANDOFF_CLAIM_TIMEOUT_MS,
      },
      {
        maxTry: 2,
        delayMs: 1500,
      }
    );
  }

  function _applyOneAiwordHandoffToken(token) {
    _setAiwordHandoffLoadingText(
      '正在认领文档到签字列表…',
      _aiwordHandoffClaimSubtext(1)
    );
    return _handoffClaimFetch(apiUrl('/api/handoff/' + encodeURIComponent(token) + '/claim-sign'))
      .then(function (claimRes) {
        var j = claimRes.data || {};
        if (j.ok && j.file && j.file.id) {
          setSaveUploadFeedback('');
          registerAiwordHandoffFile(j.file.id, j.context, j.files);
          pendingSignFiles = [];
          if (fileInput) fileInput.value = '';
          if (dirInput) dirInput.value = '';
          if (fileHint) {
            fileHint.textContent = '已从 aiword 登记文件（复用 FTP），可在下方继续签字';
          }
          if (saveBtn) saveBtn.disabled = true;
          setNeedSignActionFeedback('正在自动识别签字位并匹配编写/审核/批准人员，请稍候…');
          renderFileList();
          scheduleRefreshFileList(2000);
          return { __handoff_done: true };
        }
        return fetch(apiUrl('/api/handoff/' + encodeURIComponent(token) + '/file'), {
          credentials: 'same-origin',
        }).then(function (r) {
          var ctxHdr = (
            r.headers.get('X-Aiword-Handoff-Context') ||
            r.headers.get('x-aiword-handoff-context') ||
            ''
          ).trim();
          var parsedCtx = _decodeHandoffCtxB64(ctxHdr);
          if (!r.ok) {
            return r.text().then(function (t) {
              throw new Error(
                '交接文件获取失败（HTTP ' +
                  r.status +
                  '）' +
                  (t ? '：' + String(t).slice(0, 200) : '')
              );
            });
          }
          var cd = r.headers.get('Content-Disposition') || '';
          return r.blob().then(function (blob) {
            return { blob: blob, cd: cd, ctx: parsedCtx };
          });
        });
      })
      .then(function (o) {
        if (o && o.__handoff_done) {
          return;
        }
        if (!o || !o.blob) {
          return;
        }
        var blob = o.blob;
        var cd = o.cd;
        var name = 'document.docx';
        var m = /filename\*?=(?:UTF-8'')?["']?([^"';]+)/i.exec(cd);
        if (m) {
          try {
            name = decodeURIComponent(m[1].replace(/['"]/g, ''));
          } catch (_) {
            name = m[1];
          }
        }
        var F = window.File;
        if (!F) throw new Error('浏览器不支持 File 构造');
        var file = new F([blob], name, { type: blob.type || 'application/octet-stream' });
        mergePendingSignFiles(filterSignFiles([file]));
        updatePendingHint();
        if (!pendingSignFiles.length) {
          throw new Error('交接文件不是可签名的 .docx / .xlsx');
        }
        var form = new FormData();
        pendingSignFiles.forEach(function (f) {
          var n =
            f.webkitRelativePath && String(f.webkitRelativePath).length
              ? f.webkitRelativePath
              : f.name;
          form.append('files', f, n);
        });
        return fetchJson(apiUrl('/api/sign/upload'), { method: 'POST', body: form }).then(function (result) {
          var jj = result.data;
          if (!jj || !jj.ok) {
            throw new Error((jj && jj.error) || '保存到列表失败');
          }
          setSaveUploadFeedback('');
          var files = jj.files || [];
          var fid =
            (jj.file && jj.file.id) || (files.length ? files[files.length - 1].id : null) || null;
          var ctxObj = o.ctx && typeof o.ctx === 'object' ? o.ctx : null;
          registerAiwordHandoffFile(fid, ctxObj, files);
          pendingSignFiles = [];
          if (fileInput) fileInput.value = '';
          if (dirInput) dirInput.value = '';
          if (fileHint) {
            fileHint.textContent = '已从 aiword 传入并保存，可在下方选择文件继续签字';
          }
          if (saveBtn) saveBtn.disabled = true;
          setNeedSignActionFeedback('正在自动识别签字位并匹配编写/审核/批准人员，请稍候…');
          renderFileList();
          scheduleRefreshFileList(2000);
          return { __handoff_done: true };
        });
      });
  }

  function maybeApplyAiwordHandoffFromQuery() {
    if (!IS_FILE_SIGN_PAGE) return Promise.resolve();
    try {
      var sp = new URLSearchParams(window.location.search || '');
      if ((sp.get('from') || '').toLowerCase() !== 'aiword') {
        return Promise.resolve();
      }
      beginPageProgress('正在载入 aiword 任务…');
      try {
        document.body.classList.add('from-aiword');
      } catch (_) {}
      __aiwordPendingBatchWorkbench = true;
      var token = (sp.get('handoff_token') || '').trim();
      var tokenListRaw = (sp.get('handoff_tokens') || '').trim();
      var batchToken = (sp.get('handoff_batch_token') || '').trim();
      var tokens = [];
      if (tokenListRaw) {
        tokenListRaw.split(',').forEach(function (x) {
          var s = String(x || '').trim();
          if (s) tokens.push(s);
        });
      }
      if (!tokens.length && token) tokens.push(token);

      var chain = Promise.resolve();
      if (!tokens.length && batchToken) {
        chain = chain.then(function () {
          _setAiwordHandoffLoadingText(
            '正在登记 aiword 文档…',
            '正在登记到签字列表，请稍候，请勿关闭页面'
          );
          return _handoffClaimFetch(
            apiUrl('/api/handoff/batch/' + encodeURIComponent(batchToken) + '/claim-sign')
          ).then(function (res) {
            var j = (res && res.data) || {};
            if (!j.ok) throw new Error(j.error || '批量交接认领失败');
            var claimN = Number(j.success_count) || (Array.isArray(j.items) ? j.items.length : 0) || 0;
            _setAiwordHandoffLoadingText(
              '正在登记 aiword 文档…',
              _aiwordHandoffClaimSubtext(claimN)
            );
            var arr = Array.isArray(j.items) ? j.items : [];
            if (!arr.length) throw new Error('批量交接为空');
            setSaveUploadFeedback('');
            var batchFiles = [];
            arr.forEach(function (it) {
              var f = (it && it.file) || null;
              if (f && f.id) {
                batchFiles.push(f);
                var ctx0 = (it && it.context) || {};
                if (ctx0 && typeof ctx0 === 'object') {
                  __aiwordHandoffCtxByFileId[String(f.id)] = ctx0;
                }
              }
            });
            if (batchFiles.length) {
              savedFiles = normalizeSavedFileRecords(batchFiles);
            } else if (Array.isArray(j.files) && j.files.length) {
              savedFiles = normalizeSavedFileRecords(j.files);
            }
            pendingSignFiles = [];
            if (fileInput) fileInput.value = '';
            if (dirInput) dirInput.value = '';
            if (fileHint) {
              fileHint.textContent = '已从 aiword 批量登记文件';
            }
            if (saveBtn) saveBtn.disabled = true;
            if (j.failure_count > 0) {
              setBatchWorkbenchMsg(
                '批量认领：成功 ' + j.success_count + '，失败 ' + j.failure_count,
                true
              );
            }
            return onAiwordHandoffFilesReady().then(function () {
              scheduleRefreshFileList(2000);
            });
          });
        });
      }

      chain = chain.then(function () {
        if (!tokens.length) return;
        var seq = Promise.resolve();
        tokens.forEach(function (tk) {
          seq = seq.then(function () {
            return _applyOneAiwordHandoffToken(tk);
          });
        });
        return seq.then(function () {
          return onAiwordHandoffFilesReady().then(function () {
            scheduleRefreshFileList(2000);
          });
        });
      });
      return chain
        .then(function () {
          try {
            if (window.history && window.history.replaceState) {
              window.history.replaceState(null, '', window.location.pathname + window.location.hash);
            }
          } catch (_) {}
        })
        .catch(function (e) {
          clearAiwordHandoffState();
          return Promise.reject(e);
        })
        .finally(function () {
          endPageProgress();
        });
    } catch (e) {
      endPageProgress();
      clearAiwordHandoffState();
      return Promise.reject(e);
    }
  }

  try {
    buildUI();
    var spBoot = new URLSearchParams(window.location.search || '');
    var deferRefreshForHandoff =
      IS_FILE_SIGN_PAGE &&
      (spBoot.get('from') || '').toLowerCase() === 'aiword' &&
      (!!(spBoot.get('handoff_token') || '').trim() ||
        !!(spBoot.get('handoff_tokens') || '').trim() ||
        !!(spBoot.get('handoff_batch_token') || '').trim());

    function _bootAfterSigners() {
      if (IS_MATERIALS_PAGE) {
        updateSubmitState();
        renderLibLocaleQuickPick();
      }
      return Promise.resolve();
    }

    // 启动时拉一次运行时配置（detect 超时等）；失败用默认值。
    // 不阻塞主链：在第一个 detect 调用前 90% 概率已就绪。
    refreshRuntimeConfig();

    var bootChain;
    if (deferRefreshForHandoff) {
      // aiword 批量去签字：先完成交接认领（重），再拉签署人列表，避免 45s 内排队超时。
      // 注意：maybeApplyAiwordHandoffFromQuery → onAiwordHandoffFilesReady →
      // processBatchWorkbenchFileIds 内部已会按需 refreshSigners，
      // 这里只在 boot 完成后兜底一次（signersList 仍为空才补刷）。
      bootChain = maybeApplyAiwordHandoffFromQuery().then(function () {
        var needRefresh =
          (IS_FILE_SIGN_PAGE || IS_MATERIALS_PAGE) &&
          (!signersList || !signersList.length);
        return (needRefresh ? refreshSigners() : Promise.resolve()).then(
          _bootAfterSigners
        );
      });
    } else {
      bootChain = Promise.resolve()
        .then(function () {
          if (!IS_FILE_SIGN_PAGE && !IS_MATERIALS_PAGE) return;
          beginPageProgress('正在加载页面…');
          var loaders = [];
          if (IS_FILE_SIGN_PAGE || IS_MATERIALS_PAGE) {
            loaders.push(
              refreshSigners({
                skipPageProgress: true,
                compact: IS_FILE_SIGN_PAGE,
                deferRender: IS_FILE_SIGN_PAGE,
              })
            );
          }
          if (IS_FILE_SIGN_PAGE) {
            loaders.push(
              refreshFileList({ softFail: true, skipPageProgress: true, silent: true })
            );
          }
          return Promise.all(loaders);
        })
        .then(function () {
          if (IS_FILE_SIGN_PAGE) {
            if (savedFiles.length) {
              enableBatchWorkbenchMode();
              mergeBatchWorkbenchRowsFromSavedFiles();
              syncHiddenBatchPicks();
              _refreshWorkbenchSlotFilterOptions();
              renderBatchWorkbenchTable();
              setBatchWorkbenchMsg('文件列表已就绪，正在后台恢复识别缓存…', false);
              hydrateFileCachesFromServer({
                onProgress: function (done, total) {
                  if (total > 0 && done < total && done % 24 === 0) {
                    setBatchWorkbenchMsg(
                      '正在恢复识别缓存 ' + done + '/' + total + '…',
                      false
                    );
                  }
                },
              })
                .then(function () {
                  setBatchWorkbenchMsg('', false);
                })
                .catch(function () {
                  setBatchWorkbenchMsg('', false);
                });
            } else {
              syncHiddenBatchPicks();
              renderBatchWorkbenchTable();
            }
            refreshSignedList({ skipPageProgress: true });
            syncLibraryRolesModeRow();
            updateSubmitState();
          }
          endPageProgress();
          return _bootAfterSigners();
        })
        .catch(function () {
          endPageProgress();
        });
    }

    bootChain
      .then(
        function () {
          if (IS_FILE_SIGN_PAGE && deferRefreshForHandoff) {
            refreshSignedList();
            syncLibraryRolesModeRow();
            updateSubmitState();
          }
          try {
            window.__SIGN_PAGE_BOOT_OK = true;
          } catch (_) {}
        },
        function (e) {
          if (IS_FILE_SIGN_PAGE && deferRefreshForHandoff) {
            refreshFileList({ softFail: true });
            refreshSignedList();
            syncLibraryRolesModeRow();
            updateSubmitState();
          } else if (IS_FILE_SIGN_PAGE) {
            refreshFileList({ softFail: true, skipPageProgress: true }).catch(function () {});
          }
          try {
            var msg = e && e.message ? e.message : String(e);
            var b = document.getElementById('signBootstrapBanner');
            if (b) {
              b.style.display = 'block';
              b.textContent = deferRefreshForHandoff
                ? 'aiword 交接失败：' + msg
                : '签字页加载异常：' + msg + '（非 aiword 入口可忽略交接提示，请确认服务已启动后点「刷新列表」）';
            }
            if (typeof setSaveUploadFeedback === 'function') {
              setSaveUploadFeedback(
                deferRefreshForHandoff ? msg || '交接失败' : '加载异常：' + msg
              );
            }
          } catch (_) {}
          try {
            window.__SIGN_PAGE_BOOT_OK = true;
          } catch (_) {}
        }
      );
  } catch (bootEx) {
    var _em = bootEx && bootEx.message ? bootEx.message : String(bootEx);
    window.__SIGN_PAGE_BOOT_FAIL_MSG = '脚本运行异常：' + _em;
    try {
      window.__SIGN_PAGE_BOOT_HALTED = true;
    } catch (_) {}
    try {
      var _ban2 = document.getElementById('signBootstrapBanner');
      if (_ban2) {
        _ban2.style.display = 'block';
        _ban2.textContent = window.__SIGN_PAGE_BOOT_FAIL_MSG;
      }
      var ph2 = document.getElementById('signerLibHint');
      if (ph2) ph2.textContent = window.__SIGN_PAGE_BOOT_FAIL_MSG;
    } catch (_) {}
  }
})();
