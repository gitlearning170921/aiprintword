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
    function restoreBtn() {
      btn.removeAttribute('aria-busy');
      btn.innerHTML = prevHtml;
      if (!opt.skipRestoreDisabled) {
        btn.disabled = wasDisabled;
      }
    }
    return Promise.resolve().then(fn).then(restoreBtn, restoreBtn);
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
  var batchSelectAll = document.getElementById('batchSelectAll');
  var batchModeCb = document.getElementById('batchModeCb');
  var batchResultMsg = document.getElementById('batchResultMsg');
  var needSignActionMsg = document.getElementById('needSignActionMsg');
  var saveUploadFeedback = document.getElementById('saveUploadFeedback');
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

  function setSaveUploadFeedback(s) {
    if (!saveUploadFeedback) return;
    if (!s) {
      saveUploadFeedback.style.display = 'none';
      saveUploadFeedback.textContent = '';
      saveUploadFeedback.className = 'btn-inline-feedback is-error';
      return;
    }
    clearFileRegionErr();
    if (signerErrMsg) {
      signerErrMsg.style.display = 'none';
      signerErrMsg.textContent = '';
    }
    saveUploadFeedback.style.display = 'block';
    saveUploadFeedback.textContent = s;
    saveUploadFeedback.className = 'btn-inline-feedback is-error';
  }

  function setRoleRowFeedback(rid, s) {
    var el = document.getElementById('needSignRowMsg_' + rid);
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
    el.className = 'btn-inline-feedback is-error';
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

  function cacheMarkDetected(fileId) {
    if (!fileId) return;
    fileUiCache[fileId] = fileUiCache[fileId] || {};
    fileUiCache[fileId].detectedOnce = true;
    if (lastDetectData && lastDetectData.ok) {
      fileUiCache[fileId].lastDetectData = lastDetectData;
      fileUiCache[fileId].lastDetectError = '';
    } else if (lastDetectError) {
      fileUiCache[fileId].lastDetectError = lastDetectError;
    }
    fileUiCache[fileId].checkedRoleIds = selectedRoleIds();
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
    syncCurrentSignerBanners();
  }

  function syncCurrentSignerBanners() {
    var sid = libSignerSelect.value || '';
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
    // stroke_sets（成对）可视为同时拥有 sig+date
    try {
      var sets = signer.stroke_sets || [];
      if (Array.isArray(sets) && sets.length) {
        for (var i = 0; i < sets.length; i++) {
          var st = sets[i] || {};
          if ((st.locale || 'zh') === loc) return true;
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
    var sid = libSignerSelect.value;
    if (!sid) {
      setLibStrokeFeedback('请先在「当前签署人」中选择一位', true);
      return Promise.reject(new Error('no_signer'));
    }
    setLibStrokeFeedback('', false);
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
            drawUrlToCanvas('lib_sig_canvas', urlSig + ts, onLibStrokeImgErr),
            drawUrlToCanvas('lib_date_canvas', urlDate + ts, onLibStrokeImgErr),
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
    batchResultMsg.textContent = text || '';
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
          if (signerListEl) signerListEl.innerHTML = '';
          if (signerLibHint) signerLibHint.textContent =
            '签署人列表加载失败：' + (j.error || '请确认服务已重启。');
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
        renderSignerLib();
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
      .catch(function (e) {
        signersList = [];
        if (signerListEl) signerListEl.innerHTML = '';
        if (signerLibHint) signerLibHint.textContent =
          '无法加载签署人列表：' + (e && e.message ? e.message : String(e));
        if (IS_FILE_SIGN_PAGE) {
          renderNeedSignTable();
          refreshRoleSaveSignerControls();
          updateBatchUi();
          updateSubmitState();
        }
        if (IS_MATERIALS_PAGE) {
          refreshStrokeItemList();
        }
      });
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
    return out;
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
        var m = Object.assign({}, currentRoleMap);
        var p = m[rid] && typeof m[rid] === 'object' ? Object.assign({}, m[rid]) : {};
        p.sig = sigSel.value || null;
        if (!roleMapEntryNonEmpty(p)) delete m[rid];
        else m[rid] = p;
        fetchJson(apiUrl('/api/sign/files/' + selectedFileId + '/role-map'), {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ map: m }),
        })
          .then(function (r) {
            var jj = r.data;
            if (!jj.ok) {
              setRoleRowFeedback(rid, jj.error || '保存映射失败');
              return;
            }
            currentRoleMap = jj.map || m;
            cachePatchCurrentRoleMap(selectedFileId, currentRoleMap);
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
            fq.date = sid2;
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
        var m = Object.assign({}, currentRoleMap);
        var p = m[rid] && typeof m[rid] === 'object' ? Object.assign({}, m[rid]) : {};
        p.date = dateSel.value || null;
        if (!roleMapEntryNonEmpty(p)) delete m[rid];
        else m[rid] = p;
        fetchJson(apiUrl('/api/sign/files/' + selectedFileId + '/role-map'), {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ map: m }),
        })
          .then(function (r) {
            var jj = r.data;
            if (!jj.ok) {
              setRoleRowFeedback(rid, jj.error || '保存映射失败');
              return;
            }
            currentRoleMap = jj.map || m;
            cachePatchCurrentRoleMap(selectedFileId, currentRoleMap);
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
      dateIsoInp.value = pair0.date_iso || todayIsoLocal();

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
        var m = Object.assign({}, currentRoleMap);
        if (!roleMapEntryNonEmpty(p)) delete m[rid];
        else m[rid] = p;
        return fetchJson(apiUrl('/api/sign/files/' + selectedFileId + '/role-map'), {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ map: m }),
        })
          .then(function (r) {
            var jj = r.data;
            if (!jj.ok) {
              setRoleRowFeedback(rid, jj.error || '保存映射失败');
              return;
            }
            currentRoleMap = jj.map || m;
            cachePatchCurrentRoleMap(selectedFileId, currentRoleMap);
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
              pvMsg.textContent = msg;
            });
          });
        }).catch(function (e) {
          pvMsg.style.display = 'block';
          pvMsg.style.color = 'var(--error)';
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
          p.date_iso = dateIsoInp.value || todayIsoLocal();
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
    var n = document.querySelectorAll('.batch-pick:checked').length;
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
    syncLibStrokeSetSelect();
    syncCurrentSignerBanners();
  });

  if (libSignerFilter) libSignerFilter.addEventListener('input', function () {
    showSignerErr('');
    syncLibSignerSelect();
    syncCurrentSignerBanners();
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
    if (!libSignerSelect.value) {
      setLibStrokeFeedback('请先在「当前签署人」中选择一位', true);
      return;
    }
    withButtonBusy(libLoadStrokesBtn, '载入中…', function () {
      return loadLibStrokesFromServer().catch(function (e) {
        if (e && e.message === 'no_signer') return;
        setLibStrokeFeedback(e.message || String(e), true);
      });
    });
  });

  if (libSaveStrokesBtn) libSaveStrokesBtn.addEventListener('click', function () {
    var sid = libSignerSelect.value;
    if (!sid) {
      setLibStrokeFeedback('请先选择签署人', true);
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
      setLibStrokeFeedback(
        '请先在上方「签名」或「日期」画布中至少手写一项（可只签一项，也可两项都签）',
        true
      );
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
          setLibStrokeFeedback(jj.error || '保存失败', true);
          return;
        }
        var nid = jj.stroke_set_id;
        setLibStrokeFeedback(
          '手写图已保存到「' +
            (libSignerSelect.options[libSignerSelect.selectedIndex].text || '') +
            '」' +
            (jj.overwritten ? '（已覆盖同内容的一套）' : '') +
            '。',
          false
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
      setLibStrokeFeedback(e.message || String(e), true);
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

  function _doBatchSignFromSubmit() {
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
    var source = signSourceValue();
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
    var payload = { file_ids: ids, source: source };
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
    submitBtn.disabled = true;
    submitBtn.innerHTML = '<span class="spinner"></span> 批量处理中…';
    fetchJson(apiUrl('/api/sign/batch'), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    })
      .then(function (r) {
        var j = r.data || {};
        if (!j.ok) {
          showErr(j.error || '批量失败');
          return;
        }
        var res = j.results || [];
        var okn = res.filter(function (x) {
          return x.ok;
        }).length;
        var lines = [];
        lines.push('批量完成：成功 ' + okn + ' / ' + res.length + '。');
        res.forEach(function (it) {
          if (!it) return;
          var nm = it.name || it.file_id || '';
          var base = (it.ok ? '✅ ' : '❌ ') + nm;
          if (!it.ok) {
            lines.push(base + '：' + (it.error || '失败'));
            return;
          }
          var apN = it.applied_n || 0;
          var skN = it.skipped_n || 0;
          var det = '';
          if (apN) det += '已签 ' + apN + ' 项' + (it.applied ? '（' + it.applied + '）' : '');
          if (skN) det += (det ? '；' : '') + '跳过 ' + skN + ' 项' + (it.skipped ? '（' + it.skipped + '）' : '');
          lines.push(base + (det ? '：' + det : ''));
        });
        showBatchResult(lines.join('\n'), false);
        refreshSignedList();
      })
      .catch(function (e) {
        showErr(e.message || String(e));
      })
      .then(function () {
        submitBtn.disabled = false;
        updateSubmitState();
        updateBatchUi();
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
                  if (newId && selectedFileId) {
                    var m = Object.assign({}, currentRoleMap);
                    var p = m[rid] && typeof m[rid] === 'object' ? Object.assign({}, m[rid]) : {};
                    if (kind === 'date') p.date = newId;
                    else p.sig = newId;
                    m[rid] = p;
                    return fetchJson(apiUrl('/api/sign/files/' + selectedFileId + '/role-map'), {
                      method: 'PUT',
                      headers: { 'Content-Type': 'application/json' },
                      body: JSON.stringify({ map: m }),
                    }).then(function (r2) {
                      var j2 = r2.data;
                      if (j2 && j2.ok) currentRoleMap = j2.map || m;
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
        if (!libSignerSelect.value) {
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
        if (canvases['piece_canvas'] && canvases['piece_canvas'].clear) canvases['piece_canvas'].clear();
        batchStep += 1;
        clearFileRegionErr();
        if (batchStep >= batchOrder.length) {
          batchActive = false;
          pieceBatchNextBtn.disabled = true;
          pieceBatchUploadBtn.disabled = false;
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
        var sid = libSignerSelect.value;
        if (!sid) {
          setPieceBatchFeedback(
            '请先在上方选择「当前签署人」（当前：' +
              (currentSignerName.textContent || '—') +
              '）',
            true
          );
          return;
        }
        if (!batchQueue.length) {
          setPieceBatchFeedback('没有可上传的队列（请先完成批量录入）', true);
          return;
        }
        function uploadPieces(overwrite) {
          return fetchJson(apiUrl('/api/sign/signers/' + sid + '/stroke-pieces'), {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ items: batchQueue, overwrite: !!overwrite }),
          });
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
              if (r && (r.error_code === 'exists' || /存在/.test(String(r.error || '')))) {
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
            var yes = window.confirm('检测到该签署人已存在同名元件，是否覆盖？');
            if (!yes) {
              setPieceBatchFeedback('已取消覆盖：队列未清空，你可以调整选择或换签署人后再上传。', true);
              return Promise.resolve();
            }
            return uploadPieces(true).then(handleUploadResult);
          }

          if (okn === 0 && (failn > 0 || batchQueue.length > 0)) {
            setPieceBatchFeedback(
              '全部未保存成功' +
                (firstErrs.length ? '（' + firstErrs.join('；') + '）' : '') +
                '。请根据提示修正后仍可从队列重试上传。',
              true
            );
            setPieceHintFeedback(
              '服务器明细：成功 0 条' + (failn ? '，失败 ' + failn + ' 条' : '') + '。',
              false
            );
            return refreshSigners();
          }
          clearFileRegionErr();
          setPieceHintFeedback(
            '批量保存完成：成功 ' + okn + ' 条' + (failn ? '，失败 ' + failn + ' 条' : '') + '。',
            false
          );
          resetPieceBatchUi();
          return refreshSigners();
        }

        withButtonBusy(pieceBatchUploadBtn, '上传中…', function () {
          return uploadPieces(false).then(handleUploadResult);
        }).catch(function (e) {
          setPieceBatchFeedback(e.message || String(e), true);
        });
      });

      pieceBatchCancelBtn.addEventListener('click', function () {
        resetPieceBatchUi();
        clearFileRegionErr();
        setPieceHintFeedback('', false);
      });
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
    var batchMode = !!(batchModeCb && batchModeCb.checked);
    var picked = document.querySelectorAll('.batch-pick:checked').length;
    var src = signSourceValue();
    var libRestrict = libraryRolesRestrictedToChecks();
    var ok = false;
    if (batchMode) {
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
    submitBtn.textContent = batchMode ? '批量生成已签名文档' : '生成已签名文档';
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
    if (!IS_FILE_SIGN_PAGE || !fileListEl || !listHint || !needSignTable) return;
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
          return fetchJson(apiUrl('/api/sign/files/' + rec.id), { method: 'DELETE' }).then(
            function (result) {
              var j = result.data;
              if (!j.ok) {
                setFileListActionFeedback(j.error || '删除失败', true);
                return;
              }
              setFileListActionFeedback('', false);
              savedFiles = j.files || [];
              if (selectedFileId === rec.id) {
                selectedFileId = null;
              }
              renderFileList();
            }
          );
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
      selectedFileId = savedFiles[0].id;
      renderFileList();
      return;
    }
    updateSubmitState();
    updateBatchUi();
    var sid = selectedFileId;
    // 首次进入页面：仅首次选中时自动识别；若曾识别过（缓存里有标记），不再自动识别
    if (sid) {
      var st = fileUiCache[sid] || {};
      if (!st.detectedOnce) {
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

  function refreshSignedList() {
    if (!IS_FILE_SIGN_PAGE || !signedListEl || !signedHint) return;
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

  function refreshStrokeItemList() {
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
    fetchJson(u)
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
      });
  }

  function refreshFileList() {
    if (!IS_FILE_SIGN_PAGE || !fileListEl || !listHint) return;
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
        // 标记：该文件已执行过识别（不论成功/失败），后续切换不再自动识别
        cacheMarkDetected(fileId);
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
            cachePatchCurrentRoleMap(fileId, currentRoleMap);
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
      .then(function () {
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
      manualRedetectNeedSignRoles();
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

  function _checkedBatchFileIds() {
    return Array.from(document.querySelectorAll('.batch-pick:checked')).map(function (el) {
      return el.getAttribute('data-id');
    }).filter(Boolean);
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
    var m0 = _deepCloneJsonish(currentRoleMap || {});
    if (!Object.keys(m0).length) {
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
    chain.then(function () {
      var okMsg = '已将“' + srcName + '”文件角色配置批量映射到 ' + okN + ' 个文件。';
      // 只给“映射目标文件”写入提示；其它文件不提示
      ids.forEach(function (fid2) {
        if (!fid2) return;
        // 如果该文件失败了，不写成功提示
        var isFailed = fail.some(function (x) {
          return String(x || '').indexOf(String(fid2) + '：') === 0;
        });
        if (!isFailed) cacheSetNeedSignNotice(fid2, okMsg, false);
      });
      // 若当前选中的不是映射目标（极少：用户切走了），则不显示成功提示
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
      // 若当前选中文档也在列表里，刷新表格显示
      renderNeedSignTable();
      updateSubmitState();
    }).then(function () {
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
    clearFileRegionErr();
    setSaveUploadFeedback('');
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
            setSaveUploadFeedback(j.error || '保存失败');
            return;
          }
          setSaveUploadFeedback('');
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
        setSaveUploadFeedback(e.message || String(e));
      })
      .then(function () {
        saveBtn.disabled = !pendingSignFiles.length;
      });
  });

  if (submitBtn) submitBtn.addEventListener('click', function () {
    showErr('');
    var batchMode = !!(batchModeCb && batchModeCb.checked);
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
              var apN = sum.applied_n || 0;
              var skN = sum.skipped_n || 0;
              var apTxt2 = sum.applied || '';
              var skTxt2 = sum.skipped || '';
              showBatchResult(
                '签字结果：' +
                  (apN ? '已签 ' + apN + ' 项' + (apTxt2 ? '（' + apTxt2 + '）' : '') + '。' : '未插入任何签名/日期（全部跳过）。') +
                  (skN ? ' 跳过 ' + skN + ' 项' + (skTxt2 ? '（' + skTxt2 + '）' : '') + '。' : ''),
                false
              );
            } else {
              // 兜底：至少提示“已生成”
              showBatchResult('已生成文档。若部分角色未选择素材，将自动跳过不处理。', false);
            }
          } catch (_) {
            showBatchResult('已生成文档。若部分角色未选择素材，将自动跳过不处理。', false);
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

  try {
    buildUI();
    // 两页复用：按容器存在决定拉取哪些数据与同步哪些 UI
    if (IS_FILE_SIGN_PAGE || IS_MATERIALS_PAGE) {
      // 文件签名页和素材页都需要「签署人/素材」数据（文件页用于 role→素材下拉选择）
      refreshSigners();
    }
    if (IS_FILE_SIGN_PAGE) {
      refreshFileList();
      refreshSignedList();
      syncLibraryRolesModeRow();
      updateSubmitState();
    }
    if (IS_MATERIALS_PAGE) {
      // 素材页：仅同步本页相关的按钮状态
      updateSubmitState();
    }
    try {
      window.__SIGN_PAGE_BOOT_OK = true;
    } catch (_) {}
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
