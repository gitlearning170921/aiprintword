/* 在线签名页逻辑；由 sign.html 以 <script src> 引入，勿在聊天里复制粘贴覆盖 */
(function () {
  'use strict';

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

  function fetchJson(url, options) {
    var opts = Object.assign({ credentials: 'include' }, options || {});
    return fetch(url, opts).then(function (res) {
      return res.text().then(function (text) {
        var t = text.trim();
        if (t.charAt(0) === '<') {
          var hint =
            res.status === 404
              ? '接口不存在（请确认已保存最新代码并重启 python app.py）。'
              : '服务器返回了网页而不是 JSON。';
          throw new Error(hint + ' HTTP ' + res.status);
        }
        try {
          return { res: res, data: JSON.parse(text) };
        } catch (e) {
          throw new Error('接口返回无法解析为 JSON：' + t.slice(0, 160));
        }
      });
    });
  }

  /**
   * 异步操作期间禁用按钮并显示 spinner，避免用户以为未点击。
   * @param {Object} opt 若 opt.skipRestoreDisabled，结束时不再恢复 disabled（由业务在 finally 里设置）
   */
  function withButtonBusy(btn, busyLabel, fn, opt) {
    opt = opt || {};
    if (!btn) return Promise.resolve(fn());
    if (btn.getAttribute('aria-busy') === 'true') return Promise.resolve();
    var prevHtml = btn.innerHTML;
    var wasDisabled = btn.disabled;
    btn.setAttribute('aria-busy', 'true');
    btn.disabled = true;
    btn.innerHTML =
      '<span class="spinner" aria-hidden="true"></span> ' + (busyLabel || '处理中…');
    return Promise.resolve()
      .then(fn)
      .finally(function () {
        btn.removeAttribute('aria-busy');
        btn.innerHTML = prevHtml;
        if (!opt.skipRestoreDisabled) {
          btn.disabled = wasDisabled;
        }
      });
  }

  if (typeof window !== 'undefined' && window.location.protocol === 'file:') {
    var w = document.getElementById('fileProtoWarn');
    if (w) w.style.display = 'block';
  }

  var fileInput = document.getElementById('fileInput');
  var dirInput = document.getElementById('dirInput');
  var fileHint = document.getElementById('fileHint');
  var saveBtn = document.getElementById('saveBtn');
  var fileListEl = document.getElementById('fileList');
  var listHint = document.getElementById('listHint');
  var roleChecks = document.getElementById('roleChecks');
  var rolePanels = document.getElementById('rolePanels');
  var submitBtn = document.getElementById('submitBtn');
  var errMsg = document.getElementById('errMsg');
  var signedListEl = document.getElementById('signedList');
  var signedHint = document.getElementById('signedHint');
  var signerLibHint = document.getElementById('signerLibHint');
  var newSignerName = document.getElementById('newSignerName');
  var addSignerBtn = document.getElementById('addSignerBtn');
  var signerListEl = document.getElementById('signerListEl');
  var signerPagerInfo = document.getElementById('signerPagerInfo');
  var signerPrevBtn = document.getElementById('signerPrevBtn');
  var signerNextBtn = document.getElementById('signerNextBtn');
  var needSignTable = document.getElementById('needSignTable');
  var redetectRolesBtn = document.getElementById('redetectRolesBtn');
  var batchSelectAll = document.getElementById('batchSelectAll');
  var batchSignBtn = document.getElementById('batchSignBtn');
  var batchResultMsg = document.getElementById('batchResultMsg');
  var signSourceMode = document.getElementById('signSourceMode');
  var signerErrMsg = document.getElementById('signerErrMsg');
  var libSignerSelect = document.getElementById('libSignerSelect');
  var libSignerFilter = document.getElementById('libSignerFilter');
  var libStrokeSetSelect = document.getElementById('libStrokeSetSelect');
  var libLocaleSelect = document.getElementById('libLocaleSelect');
  var libClearSigBtn = document.getElementById('libClearSigBtn');
  var libClearDateBtn = document.getElementById('libClearDateBtn');
  var libLoadStrokesBtn = document.getElementById('libLoadStrokesBtn');
  var libSaveStrokesBtn = document.getElementById('libSaveStrokesBtn');
  var strokeItemsHint = document.getElementById('strokeItemsHint');
  var strokeItemSearchInput = document.getElementById('strokeItemSearchInput');
  var strokeItemSearchBtn = document.getElementById('strokeItemSearchBtn');
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
  var btnRefreshSigned = document.getElementById('btnRefreshSigned');

  if (
    !fileInput ||
    !fileHint ||
    !saveBtn ||
    !fileListEl ||
    !listHint ||
    !roleChecks ||
    !rolePanels ||
    !submitBtn ||
    !errMsg ||
    !signedListEl ||
    !signedHint ||
    !signerLibHint ||
    !newSignerName ||
    !addSignerBtn ||
    !signerListEl ||
    !signerPagerInfo ||
    !signerPrevBtn ||
    !signerNextBtn ||
    !needSignTable ||
    !redetectRolesBtn ||
    !batchSelectAll ||
    !batchSignBtn ||
    !batchResultMsg ||
    !signSourceMode ||
    !signerErrMsg ||
    !libSignerSelect ||
    !libSignerFilter ||
    !libStrokeSetSelect ||
    !libLocaleSelect ||
    !libClearSigBtn ||
    !libClearDateBtn ||
    !libLoadStrokesBtn ||
    !libSaveStrokesBtn ||
    !strokeItemsHint ||
    !strokeItemSearchInput ||
    !strokeItemSearchBtn ||
    !strokeItemPagerInfo ||
    !strokeItemPrevBtn ||
    !strokeItemNextBtn ||
    !strokeItemListEl ||
    !signedSearchRow ||
    !signedSearchInput ||
    !signedSearchBtn ||
    !signedPagerInfo ||
    !signedPrevBtn ||
    !signedNextBtn
  ) {
    return;
  }

  var redetectRolesBtnDefaultHtml = redetectRolesBtn.innerHTML;

  var canvases = {};
  var selectedFileId = null;
  var savedFiles = [];
  var pendingSignFiles = [];
  var lastDetectData = null;
  /** 与 lastDetectData 对应的 file_id，用于避免列表重绘时误判“已检测”而跳过 /api/sign/detect */
  var lastDetectFileId = null;
  /** 最近一次 /api/sign/detect 失败时的错误文案（用于提示，不阻塞手动勾选角色） */
  var lastDetectError = '';
  /** 正在请求检测的 file_id，避免同一文件并发重复 detect */
  var detectInFlightFor = null;
  /** 每次发起 detect 自增，用于 finally 中只恢复「最后一次」请求的 UI 状态 */
  var detectRequestSeq = 0;
  /** 交错请求时仅「当前这一轮」结束才清除 detectInFlightFor，避免长时间运行后按钮/状态卡死 */
  var detectEpoch = 0;
  var currentRoleMap = {};
  var signersList = [];
  var signersDbShare = false;

  var signerPageIndex = 0;
  var signerPageSize = 3;
  var signedListPage = 1;
  var signedListPageSize = 10;
  var signedListQ = '';
  var strokeItemPage = 1;
  // 已存储签字图片：默认 3 条/页（可翻页）
  var strokeItemPageSize = 3;
  var strokeItemQ = '';
  var roleLocaleMap = {};
  var roleLocaleManual = {};
  // 本文件角色映射：每个角色的素材下拉框独立筛选（按签署人姓名/ID）
  var roleItemFilterQ = {};

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
    if (s) {
      errMsg.style.display = 'none';
      errMsg.textContent = '';
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

  function syncLibSignerSelect() {
    var prev = libSignerSelect.value;
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
    if (prev && signersList.some(function (x) {
      return x.id === prev;
    })) {
      libSignerSelect.value = prev;
    }
    syncLibStrokeSetSelect();
  }

  function syncLibStrokeSetSelect() {
    var prev = libStrokeSetSelect.value;
    libStrokeSetSelect.innerHTML = '';
    var o0 = document.createElement('option');
    o0.value = '';
    o0.textContent = '不指定成对套（载入该人最近一套签名+日期）';
    libStrokeSetSelect.appendChild(o0);
    var sid = libSignerSelect.value;
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

  function loadLibStrokesFromServer() {
    var sid = libSignerSelect.value;
    if (!sid) {
      showSignerErr('请先在「当前签署人」中选择一位');
      return Promise.reject(new Error('no_signer'));
    }
    showSignerErr('');
    var ts = '?t=' + Date.now();
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
            drawUrlToCanvas('lib_sig_canvas', urlSig + ts, showSignerErr),
            drawUrlToCanvas('lib_date_canvas', urlDate + ts, showSignerErr),
          ])
            .then(function () {
              resolve();
            })
            .catch(reject);
        });
      });
    });
  }

  function showSignerListLoading() {
    signerListEl.innerHTML = '';
    var li = document.createElement('li');
    li.textContent = '正在加载签署人列表…';
    li.style.border = 'none';
    li.style.background = 'transparent';
    signerListEl.appendChild(li);
  }

  function showFileListLoading() {
    fileListEl.innerHTML = '';
    listHint.style.display = 'block';
    listHint.textContent = '正在加载文件列表…';
  }

  function showSignedListLoading() {
    signedListEl.innerHTML = '';
    signedHint.style.display = 'block';
    signedHint.textContent = '正在加载已签名列表…';
  }

  function showStrokeItemsLoading() {
    strokeItemListEl.innerHTML = '';
    strokeItemsHint.style.display = 'block';
    strokeItemsHint.textContent = '正在加载已存储签字图片…';
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
    return e === '.docx' || e === '.xlsx';
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
        '可多次选择并累加 .docx / .xlsx，或使用「选择文件夹」上传整目录（其他类型会自动忽略）';
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
    var targetW = kind === 'date' ? 900 : 1200;
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
      octx.clearRect(0, 0, targetW, targetH);
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
    batchResultMsg.textContent = text || '';
  }

  function drawUrlToCanvas(canvasId, url, onImgErr) {
    var canvas = document.getElementById(canvasId);
    if (!canvas) return Promise.resolve(false);
    var ctx = canvas.getContext('2d');
    return new Promise(function (resolve) {
      var img = new Image();
      img.crossOrigin = 'anonymous';
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
        var msg = '无法加载手写图（请确认该签署人已保存对应签名/日期）';
        if (typeof onImgErr === 'function') {
          onImgErr(msg);
        } else {
          showErr(msg);
        }
        resolve(false);
      };
      img.src = url;
    });
  }

  function refreshSigners() {
    showSignerListLoading();
    return fetchJson(apiUrl('/api/sign/signers') + '?_=' + Date.now(), {
      cache: 'no-store',
    })
      .then(function (result) {
        var j = result.data || {};
        if (!j.ok) {
          signersList = [];
          signerListEl.innerHTML = '';
          signerLibHint.textContent =
            '签署人列表加载失败：' + (j.error || '请确认服务已重启。');
          renderNeedSignTable();
          updateBatchUi();
          updateSubmitState();
          return;
        }
        signersDbShare = !!j.db_share;
        signersList = Array.isArray(j.signers) ? j.signers : [];
        signerLibHint.textContent = signersDbShare
          ? '已启用 MySQL：签署人笔迹可在多台电脑复用。'
          : '当前为会话目录存储：笔迹仅保存在本浏览器对应会话目录，换浏览器需重新添加。';
        renderSignerLib();
        renderNeedSignTable();
        updateBatchUi();
        updateSubmitState();
        refreshStrokeItemList();
      })
      .catch(function (e) {
        signersList = [];
        signerListEl.innerHTML = '';
        signerLibHint.textContent =
          '无法加载签署人列表：' + (e && e.message ? e.message : String(e));
        renderNeedSignTable();
        updateBatchUi();
        updateSubmitState();
        refreshStrokeItemList();
      });
  }

  function renderSignerLib() {
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
    signerPagerInfo.textContent =
      '已录入 ' + total + ' 位签署人 · 第 ' + (signerPageIndex + 1) + ' / ' + pageCount + ' 页（默认仅显示 3 位）';
    signerPrevBtn.disabled = signerPageIndex <= 0;
    signerNextBtn.disabled = signerPageIndex >= pageCount - 1;
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
    return out;
  }

  function renderNeedSignTable() {
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
      td1.style.padding = '8px';
      td1.style.borderBottom = '1px solid var(--border)';
      var td2 = document.createElement('td');
      td2.style.padding = '8px';
      td2.style.borderBottom = '1px solid var(--border)';
      td2.textContent =
        row.conf != null && typeof row.conf === 'number' ? row.conf.toFixed(2) : '—';
      var td3 = document.createElement('td');
      td3.style.padding = '8px';
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
        var m = Object.assign({}, currentRoleMap);
        var p = m[rid] && typeof m[rid] === 'object' ? Object.assign({}, m[rid]) : {};
        p.sig = sigSel.value || null;
        if (!p.sig && !p.date) delete m[rid];
        else m[rid] = p;
        fetchJson(apiUrl('/api/sign/files/' + selectedFileId + '/role-map'), {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ map: m }),
        })
          .then(function (r) {
            var jj = r.data;
            if (!jj.ok) {
              showErr(jj.error || '保存映射失败');
              return;
            }
            currentRoleMap = jj.map || m;
            showErr('');
          })
          .catch(function (e) {
            showErr(e.message || String(e));
          });
      });
      td3.appendChild(sigFilter);
      td3.appendChild(sigSel);
      var td3b = document.createElement('td');
      td3b.style.padding = '8px';
      td3b.style.borderBottom = '1px solid var(--border)';
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
        var m = Object.assign({}, currentRoleMap);
        var p = m[rid] && typeof m[rid] === 'object' ? Object.assign({}, m[rid]) : {};
        p.date = dateSel.value || null;
        if (!p.sig && !p.date) delete m[rid];
        else m[rid] = p;
        fetchJson(apiUrl('/api/sign/files/' + selectedFileId + '/role-map'), {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ map: m }),
        })
          .then(function (r) {
            var jj = r.data;
            if (!jj.ok) {
              showErr(jj.error || '保存映射失败');
              return;
            }
            currentRoleMap = jj.map || m;
            showErr('');
          })
          .catch(function (e) {
            showErr(e.message || String(e));
          });
      });
      td3b.appendChild(dateFilter);
      td3b.appendChild(dateSel);
      var td4 = document.createElement('td');
      td4.style.padding = '8px';
      td4.style.borderBottom = '1px solid var(--border)';
      td4.style.whiteSpace = 'nowrap';
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
          showErr('请先选择签名素材');
          return;
        }
        withButtonBusy(bLoad, '载入中…', function () {
          return new Promise(function (resolve) {
            showErr('');
            var ts = '?t=' + Date.now();
            setRoleChecked(rid, true);
            requestAnimationFrame(function () {
              requestAnimationFrame(function () {
                resizeCanvasesForRoles([rid]);
                drawUrlToCanvas(
                  'sig_' + rid,
                  apiUrl('/api/sign/stroke-items/' + itemId + '/png') + ts
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
        var itemId = dateSel.value;
        if (!itemId) {
          showErr('请先选择日期素材');
          return;
        }
        withButtonBusy(bLoadDate, '载入中…', function () {
          return new Promise(function (resolve) {
            showErr('');
            var ts = '?t=' + Date.now();
            setRoleChecked(rid, true);
            requestAnimationFrame(function () {
              requestAnimationFrame(function () {
                resizeCanvasesForRoles([rid]);
                drawUrlToCanvas(
                  'date_' + rid,
                  apiUrl('/api/sign/stroke-items/' + itemId + '/png') + ts
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
      var bSave = document.createElement('button');
      bSave.type = 'button';
      bSave.className = 'btn btn-secondary';
      bSave.textContent = '入库签名并绑定';
      bSave.title = '将本角色签名画布写入签名库（按内容去重覆盖），并绑定本角色';
      bSave.addEventListener('click', function () {
        var signerForSave = signerIdFromStrokeSetSelect(sigSel) || signerIdFromStrokeSetSelect(dateSel);
        if (!signerForSave) {
          showErr('请先选择任意一条素材以确定保存到哪一签署人（无素材时请先在「一、」中保存）');
          return;
        }
        setRoleChecked(rid, true);
        requestAnimationFrame(function () {
          requestAnimationFrame(function () {
            resizeCanvasesForRoles([rid]);
            var sigC = document.getElementById('sig_' + rid);
            if (isCanvasBlank(sigC)) {
              showErr('请先在「' + roleLabel(rid) + '」签名画布上手写签名');
              return;
            }
            var fd = new FormData();
            fd.append('sig', _normalizedPngDataUrl(sigC, 'sig'));
            var baseSel = sigSel.value ? sigSel : (dateSel.value ? dateSel : sigSel);
            var locV = (locSel && locSel.value) ? locSel.value : (roleLocaleMap[rid] || 'auto');
            var finalLoc =
              locV === 'en' || locV === 'zh'
                ? locV
                : localeFromStrokeSetOption(baseSel) || (libLocaleSelect && libLocaleSelect.value) || 'zh';
            fd.append('locale', finalLoc);
            withButtonBusy(bSave, '保存中…', function () {
              return fetchJson(apiUrl('/api/sign/signers/' + signerForSave + '/strokes'), {
                method: 'PUT',
                body: fd,
              })
                .then(function (r) {
                  var jj = r.data;
                  if (!jj.ok) {
                    showErr(jj.error || '保存失败');
                    return;
                  }
                  showErr('');
                  var newId = jj.sig_item_id;
                  if (newId && selectedFileId) {
                    var m = Object.assign({}, currentRoleMap);
                    var p = m[rid] && typeof m[rid] === 'object' ? Object.assign({}, m[rid]) : {};
                    p.sig = newId;
                    m[rid] = p;
                    return fetchJson(apiUrl('/api/sign/files/' + selectedFileId + '/role-map'), {
                      method: 'PUT',
                      headers: { 'Content-Type': 'application/json' },
                      body: JSON.stringify({ map: m }),
                    }).then(function (r2) {
                      var j2 = r2.data;
                      if (j2 && j2.ok) currentRoleMap = j2.map || m;
                      showBatchResult('签名已入库并已绑定本角色。', false);
                      return refreshSigners();
                    });
                  }
                  showBatchResult('签名已保存到所选签署人。', false);
                  return refreshSigners();
                })
                .then(function () {
                  renderNeedSignTable();
                });
            }).catch(function (e) {
              showErr(e.message || String(e));
            });
          });
        });
      });
      var bSaveDate = document.createElement('button');
      bSaveDate.type = 'button';
      bSaveDate.className = 'btn btn-secondary';
      bSaveDate.textContent = '入库日期并绑定';
      bSaveDate.title = '将本角色日期画布写入日期库（按内容去重覆盖），并绑定本角色';
      bSaveDate.addEventListener('click', function () {
        var signerForSave = signerIdFromStrokeSetSelect(sigSel) || signerIdFromStrokeSetSelect(dateSel);
        if (!signerForSave) {
          showErr('请先选择任意一条素材以确定保存到哪一签署人（无素材时请先在「一、」中保存）');
          return;
        }
        setRoleChecked(rid, true);
        requestAnimationFrame(function () {
          requestAnimationFrame(function () {
            resizeCanvasesForRoles([rid]);
            var dateC = document.getElementById('date_' + rid);
            if (isCanvasBlank(dateC)) {
              showErr('请先在「' + roleLabel(rid) + '」日期画布上手写日期');
              return;
            }
            var fd = new FormData();
            fd.append('date', _normalizedPngDataUrl(dateC, 'date'));
            var baseSel = dateSel.value ? dateSel : (sigSel.value ? sigSel : dateSel);
            var locV = (locSel && locSel.value) ? locSel.value : (roleLocaleMap[rid] || 'auto');
            var finalLoc =
              locV === 'en' || locV === 'zh'
                ? locV
                : localeFromStrokeSetOption(baseSel) || (libLocaleSelect && libLocaleSelect.value) || 'zh';
            fd.append('locale', finalLoc);
            withButtonBusy(bSaveDate, '保存中…', function () {
              return fetchJson(apiUrl('/api/sign/signers/' + signerForSave + '/strokes'), {
                method: 'PUT',
                body: fd,
              })
                .then(function (r) {
                  var jj = r.data;
                  if (!jj.ok) {
                    showErr(jj.error || '保存失败');
                    return;
                  }
                  showErr('');
                  var newId = jj.date_item_id;
                  if (newId && selectedFileId) {
                    var m = Object.assign({}, currentRoleMap);
                    var p = m[rid] && typeof m[rid] === 'object' ? Object.assign({}, m[rid]) : {};
                    p.date = newId;
                    m[rid] = p;
                    return fetchJson(apiUrl('/api/sign/files/' + selectedFileId + '/role-map'), {
                      method: 'PUT',
                      headers: { 'Content-Type': 'application/json' },
                      body: JSON.stringify({ map: m }),
                    }).then(function (r2) {
                      var j2 = r2.data;
                      if (j2 && j2.ok) currentRoleMap = j2.map || m;
                      showBatchResult('日期已入库并已绑定本角色。', false);
                      return refreshSigners();
                    });
                  }
                  showBatchResult('日期已保存到所选签署人。', false);
                  return refreshSigners();
                })
                .then(function () {
                  renderNeedSignTable();
                });
            }).catch(function (e) {
              showErr(e.message || String(e));
            });
          });
        });
      });
      td4.appendChild(locSel);
      td4.appendChild(bLoad);
      td4.appendChild(bLoadDate);
      td4.appendChild(bSave);
      td4.appendChild(bSaveDate);
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
    var n = document.querySelectorAll('.batch-pick:checked').length;
    batchSignBtn.disabled = !signersDbShare || n === 0;
  }

  batchSelectAll.addEventListener('change', function () {
    document.querySelectorAll('.batch-pick').forEach(function (cb) {
      cb.checked = batchSelectAll.checked;
    });
    updateBatchUi();
  });

  addSignerBtn.addEventListener('click', function () {
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

  libSignerSelect.addEventListener('change', function () {
    showSignerErr('');
    syncLibStrokeSetSelect();
  });

  libSignerFilter.addEventListener('input', function () {
    showSignerErr('');
    syncLibSignerSelect();
  });

  signerPrevBtn.addEventListener('click', function () {
    signerPageIndex = Math.max(0, signerPageIndex - 1);
    renderSignerLib();
  });

  signerNextBtn.addEventListener('click', function () {
    signerPageIndex = signerPageIndex + 1;
    renderSignerLib();
  });

  libClearSigBtn.addEventListener('click', function () {
    if (canvases['lib_sig_canvas'] && canvases['lib_sig_canvas'].clear) {
      canvases['lib_sig_canvas'].clear();
    }
  });

  libClearDateBtn.addEventListener('click', function () {
    if (canvases['lib_date_canvas'] && canvases['lib_date_canvas'].clear) {
      canvases['lib_date_canvas'].clear();
    }
  });

  libLoadStrokesBtn.addEventListener('click', function () {
    if (!libSignerSelect.value) {
      showSignerErr('请先在「当前签署人」中选择一位');
      return;
    }
    withButtonBusy(libLoadStrokesBtn, '载入中…', function () {
      return loadLibStrokesFromServer().catch(function (e) {
        if (e && e.message === 'no_signer') return;
        showSignerErr(e.message || String(e));
      });
    });
  });

  libSaveStrokesBtn.addEventListener('click', function () {
    var sid = libSignerSelect.value;
    if (!sid) {
      showSignerErr('请先选择签署人');
      return;
    }
    if (canvases['lib_sig_canvas'] && canvases['lib_sig_canvas'].resize) {
      canvases['lib_sig_canvas'].resize();
    }
    if (canvases['lib_date_canvas'] && canvases['lib_date_canvas'].resize) {
      canvases['lib_date_canvas'].resize();
    }
    var sigC = document.getElementById('lib_sig_canvas');
    var dateC = document.getElementById('lib_date_canvas');
    var sigBlank = isCanvasBlank(sigC);
    var dateBlank = isCanvasBlank(dateC);
    if (sigBlank && dateBlank) {
      showSignerErr('请先在上方「签名」或「日期」画布中至少手写一项（可只签一项，也可两项都签）');
      return;
    }
    var fd = new FormData();
    if (!sigBlank) fd.append('sig', _normalizedPngDataUrl(sigC, 'sig'));
    if (!dateBlank) fd.append('date', _normalizedPngDataUrl(dateC, 'date'));
    fd.append('locale', (libLocaleSelect && libLocaleSelect.value) ? libLocaleSelect.value : 'zh');
    withButtonBusy(libSaveStrokesBtn, '保存中…', function () {
      return fetchJson(apiUrl('/api/sign/signers/' + sid + '/strokes'), {
        method: 'PUT',
        body: fd,
      }).then(function (r) {
        var jj = r.data || {};
        if (!jj.ok) {
          showSignerErr(jj.error || '保存失败');
          return;
        }
        var nid = jj.stroke_set_id;
        showSignerErr(
          '手写图已保存到「' +
            (libSignerSelect.options[libSignerSelect.selectedIndex].text || '') +
            '」' +
            (jj.overwritten ? '（已覆盖同内容的一套）' : '') +
            '。'
        );
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
      showSignerErr(e.message || String(e));
    });
  });

  if (btnRefreshSigners) {
    btnRefreshSigners.addEventListener('click', function () {
      refreshSigners();
    });
  }
  if (btnRefreshStrokeItems) {
    btnRefreshStrokeItems.addEventListener('click', function () {
      refreshStrokeItemList();
    });
  }
  if (btnRefreshFiles) {
    btnRefreshFiles.addEventListener('click', function () {
      refreshFileList();
    });
  }
  if (btnRefreshSigned) {
    btnRefreshSigned.addEventListener('click', function () {
      refreshSignedList();
    });
  }

  batchSignBtn.addEventListener('click', function () {
    var ids = Array.from(document.querySelectorAll('.batch-pick:checked')).map(function (el) {
      return el.getAttribute('data-id');
    });
    if (!ids.length) {
      showErr('请先勾选要批量签名的文件');
      return;
    }
    if (!signersDbShare) {
      showErr('批量签名需要启用 MySQL（MYSQL_HOST）');
      return;
    }
    var source = (signSourceMode.value || 'canvas').trim().toLowerCase();
    var rolesForBatch = selectedRoleIds();
    if (!rolesForBatch.length) {
      showErr('请至少勾选一个角色（用于批量签）');
      return;
    }
    var payload = { file_ids: ids, source: source, roles: rolesForBatch };
    if (source === 'canvas') {
      resizeCanvasesForRoles(rolesForBatch);
      for (var ri2 = 0; ri2 < rolesForBatch.length; ri2++) {
        var rid2 = rolesForBatch[ri2];
        var sigC2 = document.getElementById('sig_' + rid2);
        var dateC2 = document.getElementById('date_' + rid2);
        if (isCanvasBlank(sigC2) || isCanvasBlank(dateC2)) {
          var role2 = ROLES.find(function (x) {
            return x.id === rid2;
          });
          var label2 = role2 ? role2.label : rid2;
          showErr('请为「' + label2 + '」完成签名与日期手写（或切换为“库映射”）');
          return;
        }
      }
      payload.sig_map = {};
      payload.date_map = {};
      rolesForBatch.forEach(function (rid3) {
        payload.sig_map[rid3] = _normalizedPngDataUrl(document.getElementById('sig_' + rid3), 'sig');
        payload.date_map[rid3] = _normalizedPngDataUrl(document.getElementById('date_' + rid3), 'date');
      });
    }
    showErr('');
    withButtonBusy(
      batchSignBtn,
      '批量处理中…',
      function () {
        return fetchJson(apiUrl('/api/sign/batch'), {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
        }).then(function (r) {
          var j = r.data || {};
          if (!j.ok) {
            showErr(j.error || '批量失败');
            return;
          }
          var res = j.results || [];
          var okn = res.filter(function (x) {
            return x.ok;
          }).length;
          showBatchResult(
            '批量完成：成功 ' + okn + ' / ' + res.length + '。可在下方「已签名文档」下载。',
            false
          );
          refreshSignedList();
        });
      },
      { skipRestoreDisabled: true }
    )
      .catch(function (e) {
        showErr(e.message || String(e));
      })
      .finally(function () {
        updateBatchUi();
      });
  });

  function buildUI() {
    ROLES.forEach(function (r) {
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
        renderNeedSignTable();
        updateSubmitState();
      });
    });

    rolePanels.querySelectorAll('[data-clear]').forEach(function (btn) {
      btn.addEventListener('click', function () {
        var id = btn.getAttribute('data-clear');
        if (canvases[id]) canvases[id].clear();
      });
    });

    ROLES.forEach(function (r) {
      canvases['sig_' + r.id] = setupCanvas(document.getElementById('sig_' + r.id), {
        lineWidth: 4.35,
        shadowBlur: 0.7,
      });
      canvases['date_' + r.id] = setupCanvas(document.getElementById('date_' + r.id), {
        lineWidth: 3.55,
        shadowBlur: 0.55,
      });
    });

    canvases['lib_sig_canvas'] = setupCanvas(document.getElementById('lib_sig_canvas'), {
      lineWidth: 4.35,
      shadowBlur: 0.7,
    });
    canvases['lib_date_canvas'] = setupCanvas(document.getElementById('lib_date_canvas'), {
      lineWidth: 3.55,
      shadowBlur: 0.55,
    });

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
    signedPrevBtn.addEventListener('click', function () {
      if (signedListPage <= 1) return;
      signedListPage -= 1;
      refreshSignedList();
    });
    signedNextBtn.addEventListener('click', function () {
      signedListPage += 1;
      refreshSignedList();
    });

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
    strokeItemPrevBtn.addEventListener('click', function () {
      if (strokeItemPage <= 1) return;
      strokeItemPage -= 1;
      refreshStrokeItemList();
    });
    strokeItemNextBtn.addEventListener('click', function () {
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
    var ok = selectedFileId && selectedRoleIds().length > 0;
    submitBtn.disabled = !ok;
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

  function renderFileList() {
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
        selectedFileId = rec.id;
        lastDetectData = null;
        lastDetectFileId = null;
        lastDetectError = '';
        document.querySelectorAll('.file-list li').forEach(function (el) {
          el.classList.remove('selected');
        });
        li.classList.add('selected');
        detectAndAutoSelectRoles(selectedFileId);
        updateSubmitState();
      });
      delBtn.addEventListener('click', function () {
        withButtonBusy(delBtn, '删除中…', function () {
          return fetchJson(apiUrl('/api/sign/files/' + rec.id), { method: 'DELETE' }).then(
            function (result) {
              var j = result.data;
              if (!j.ok) {
                showErr(j.error || '删除失败');
                return;
              }
              savedFiles = j.files || [];
              if (selectedFileId === rec.id) {
                selectedFileId = null;
              }
              renderFileList();
            }
          );
        }).catch(function (e) {
          showErr(e.message || String(e));
        });
      });

      li.appendChild(batchCb);
      li.appendChild(radio);
      li.appendChild(lbl);
      li.appendChild(delBtn);
      fileListEl.appendChild(li);
    });
    if (!selectedFileId && savedFiles.length) {
      selectedFileId = savedFiles[0].id;
      renderFileList();
      return;
    }
    updateSubmitState();
    updateBatchUi();
    var sid = selectedFileId;
    var hasValidDetect =
      sid &&
      lastDetectData &&
      lastDetectData.ok &&
      String(lastDetectFileId) === String(sid);
    if (sid && !hasValidDetect) {
      detectAndAutoSelectRoles(sid);
    } else if (sid) {
      fetchJson(apiUrl('/api/sign/files/' + sid + '/role-map'))
        .then(function (r) {
          var jj = r.data || {};
          if (jj.ok) currentRoleMap = jj.map || {};
          renderNeedSignTable();
        })
        .catch(function () {});
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
    var items = (j && j.items) || [];
    var dbShare = !!(j && j.db_share);
    var total = typeof (j && j.total) === 'number' ? j.total : items.length;
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
    if (!items.length) {
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
    items.forEach(function (it) {
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
                showErr(jj.error || '删除失败');
                return;
              }
              refreshSignedList();
            }
          );
        }).catch(function (e) {
          showErr(e.message || String(e));
        });
      });

      li.appendChild(wrap);
      li.appendChild(dl);
      li.appendChild(delBtn);
      signedListEl.appendChild(li);
    });
  }

  function refreshSignedList() {
    showSignedListLoading();
    var u =
      apiUrl('/api/sign/signed') +
      '?q=' +
      encodeURIComponent(signedListQ) +
      '&page=' +
      signedListPage +
      '&page_size=' +
      signedListPageSize +
      '&_=' +
      Date.now();
    fetchJson(u)
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
      strokeItemPager.style.display = 'none';
      return;
    }
    strokeItemPager.style.display = '';
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
                showSignerErr(jj.error || '删除失败');
                return;
              }
              refreshStrokeItemList();
            }
          );
        }).catch(function (e) {
          showSignerErr(e.message || String(e));
        });
      });

      li.appendChild(wrap);
      li.appendChild(dl);
      li.appendChild(delBtn);
      strokeItemListEl.appendChild(li);
    });
  }

  function refreshStrokeItemList() {
    showStrokeItemsLoading();
    var u =
      apiUrl('/api/sign/stroke-items') +
      '?q=' +
      encodeURIComponent(strokeItemQ) +
      '&page=' +
      strokeItemPage +
      '&page_size=' +
      strokeItemPageSize +
      '&_=' +
      Date.now();
    fetchJson(u)
      .then(function (result) {
        var j = result.data;
        if (!j.ok) {
          strokeItemListEl.innerHTML = '';
          strokeItemsHint.style.display = 'block';
          strokeItemPager.style.display = 'none';
          strokeItemsHint.textContent =
            '签字素材列表加载失败：' + (j.error || '请稍后重试。');
          return;
        }
        renderStrokeItemList(j);
      })
      .catch(function (e) {
        strokeItemListEl.innerHTML = '';
        strokeItemsHint.style.display = 'block';
        strokeItemPager.style.display = 'none';
        strokeItemsHint.textContent =
          '签字素材列表加载失败：' + (e && e.message ? e.message : String(e));
      });
  }

  function refreshFileList() {
    showFileListLoading();
    fetchJson(apiUrl('/api/sign/files'))
      .then(function (result) {
        var j = result.data;
        if (!j.ok || !Array.isArray(j.files)) {
          savedFiles = [];
          if (selectedFileId) selectedFileId = null;
          renderFileList();
          listHint.style.display = 'block';
          listHint.textContent =
            '文件列表加载失败' + (j && j.error ? ('：' + j.error) : '。请刷新页面或确认服务已启动。');
          return;
        }
        savedFiles = j.files;
        if (selectedFileId && !savedFiles.some(function (f) {
          return f.id === selectedFileId;
        })) {
          selectedFileId = null;
        }
        renderFileList();
      })
      .catch(function (e) {
        savedFiles = [];
        if (selectedFileId) selectedFileId = null;
        renderFileList();
        listHint.style.display = 'block';
        listHint.textContent =
          '文件列表加载失败：' + (e && e.message ? e.message : String(e));
      });
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

  function detectAndAutoSelectRoles(fileId) {
    if (!fileId) return false;
    if (String(detectInFlightFor) === String(fileId)) return false;
    detectInFlightFor = fileId;
    var myEpoch = ++detectEpoch;
    var seq = ++detectRequestSeq;
    lastDetectData = null;
    lastDetectFileId = null;
    lastDetectError = '';
    resetAllRoleChecks();
    redetectRolesBtn.disabled = true;
    redetectRolesBtn.innerHTML =
      '<span class="spinner" aria-hidden="true"></span> 分析中…';
    needSignTable.innerHTML = '';
    needSignTable.textContent = '正在分析模板与角色映射…';
    fetchJson(apiUrl('/api/sign/detect?file_id=' + encodeURIComponent(fileId)))
      .then(function (result) {
        if (String(selectedFileId) !== String(fileId)) {
          return { __abort: true };
        }
        var j = result.data || {};
        if (j.ok) {
          lastDetectError = '';
          lastDetectData = j;
          lastDetectFileId = fileId;
        } else {
          lastDetectData = null;
          lastDetectFileId = null;
          lastDetectError = (j && j.error) || '识别接口返回失败';
        }
        if (j.ok) {
          var roles = mergeDetectedRolesForUi();
          if (roles.length) {
            roles.forEach(function (r) {
              if (r && r.id) setRoleChecked(r.id, true);
            });
            requestAnimationFrame(function () {
              requestAnimationFrame(function () {
                resizeCanvasesForRoles(
                  roles
                    .map(function (x) {
                      return x && x.id;
                    })
                    .filter(Boolean)
                );
              });
            });
          }
        }
        return fetchJson(apiUrl('/api/sign/files/' + fileId + '/role-map')).then(function (rm) {
          if (String(selectedFileId) !== String(fileId)) {
            return { __abort: true };
          }
          if (rm && rm.data && rm.data.ok) {
            currentRoleMap = rm.data.map || {};
          }
          return { __abort: false };
        });
      })
      .then(function (pack) {
        if (pack && pack.__abort) return;
        renderNeedSignTable();
        updateSubmitState();
      })
      .catch(function (err) {
        if (String(selectedFileId) === String(fileId)) {
          lastDetectData = null;
          lastDetectFileId = null;
          lastDetectError =
            (err && err.message) ||
            '识别请求失败（网络或服务异常）。请稍后重试「重新识别」。';
          renderNeedSignTable();
          updateSubmitState();
        }
      })
      .finally(function () {
        if (myEpoch === detectEpoch) {
          detectInFlightFor = null;
        }
        if (seq === detectRequestSeq) {
          redetectRolesBtn.disabled = false;
          redetectRolesBtn.innerHTML = redetectRolesBtnDefaultHtml;
        }
      });
    return true;
  }

  function manualRedetectNeedSignRoles() {
    showErr('');
    if (!selectedFileId) {
      showErr('请先在上方的文件列表中选择一项。');
      return;
    }
    if (String(detectInFlightFor) === String(selectedFileId)) {
      showErr('正在分析当前文件，请稍候再试。');
      return;
    }
    lastDetectData = null;
    lastDetectFileId = null;
    lastDetectError = '';
    detectAndAutoSelectRoles(selectedFileId);
  }

  redetectRolesBtn.addEventListener('click', function () {
    manualRedetectNeedSignRoles();
  });

  function showErr(s) {
    if (s) {
      signerErrMsg.style.display = 'none';
      signerErrMsg.textContent = '';
    }
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

  fileInput.addEventListener('change', function () {
    mergePendingSignFiles(filterSignFiles(Array.from(fileInput.files || [])));
    fileInput.value = '';
    updatePendingHint();
  });

  if (dirInput) {
    dirInput.addEventListener('change', function () {
      mergePendingSignFiles(filterSignFiles(Array.from(dirInput.files || [])));
      dirInput.value = '';
      updatePendingHint();
    });
  }

  saveBtn.addEventListener('click', function () {
    showErr('');
    if (!pendingSignFiles.length) return;
    var form = new FormData();
    pendingSignFiles.forEach(function (f) {
      var name =
        f.webkitRelativePath && String(f.webkitRelativePath).length
          ? f.webkitRelativePath
          : f.name;
      form.append('files', f, name);
    });
    withButtonBusy(saveBtn, '上传中…', function () {
      return fetchJson(apiUrl('/api/sign/upload'), { method: 'POST', body: form }).then(
        function (result) {
          var j = result.data;
          if (!j.ok) {
            showErr(j.error || '保存失败');
            return;
          }
          savedFiles = j.files || [];
          selectedFileId =
            (j.file && j.file.id) ||
            (savedFiles.length && savedFiles[savedFiles.length - 1].id);
          pendingSignFiles = [];
          fileInput.value = '';
          if (dirInput) dirInput.value = '';
          fileHint.textContent = '已保存，可继续添加或从下方列表选择';
          saveBtn.disabled = true;
          renderFileList();
        }
      );
    }, { skipRestoreDisabled: true })
      .catch(function (e) {
        showErr(e.message || String(e));
      })
      .finally(function () {
        saveBtn.disabled = !pendingSignFiles.length;
      });
  });

  submitBtn.addEventListener('click', function () {
    showErr('');
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
    var roles = selectedRoleIds();
    if (!roles.length) {
      showErr('请至少勾选一个角色');
      return;
    }
    var source2 = (signSourceMode.value || 'canvas').trim().toLowerCase();
    if (source2 === 'canvas') {
      resizeCanvasesForRoles(roles);
      for (var ri = 0; ri < roles.length; ri++) {
        var id = roles[ri];
        var sigC = document.getElementById('sig_' + id);
        var dateC = document.getElementById('date_' + id);
        if (isCanvasBlank(sigC) || isCanvasBlank(dateC)) {
          var role = ROLES.find(function (x) {
            return x.id === id;
          });
          var label = role ? role.label : id;
          showErr('请为「' + label + '」完成签名与日期手写（或切换为“库映射”）');
          return;
        }
      }
    }

    var form = new FormData();
    form.append('file_id', selectedFileId);
    form.append('roles', JSON.stringify(roles));
    form.append('sign_source', source2);
    if (source2 === 'canvas') {
      roles.forEach(function (id) {
        form.append('sig_' + id, _normalizedPngDataUrl(document.getElementById('sig_' + id), 'sig'));
        form.append('date_' + id, _normalizedPngDataUrl(document.getElementById('date_' + id), 'date'));
      });
    }

    submitBtn.disabled = true;
    submitBtn.innerHTML = '<span class="spinner"></span> 处理中…';

    fetch(apiUrl('/api/sign'), { method: 'POST', body: form, credentials: 'include' })
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
          refreshSignedList();
        });
      })
      .catch(function (e) {
        showErr(e.message || String(e));
      })
      .then(function () {
        submitBtn.disabled = false;
        submitBtn.textContent = '生成已签名文档';
        updateSubmitState();
      });
  });

  buildUI();
  refreshSigners();
  refreshFileList();
  refreshSignedList();
  updateSubmitState();
})();
