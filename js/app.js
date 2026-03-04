/* global pdfjsLib, XLSX */

(function () {
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

  const pdfInput = document.getElementById('pdfInput');
  const uploadZone = document.getElementById('uploadZone');
  const uploadFilenames = document.getElementById('uploadFilenames');
  const btnExtract = document.getElementById('btnExtract');
  const extractStatus = document.getElementById('extractStatus');
  const loadingWrap = document.getElementById('loadingWrap');
  const loadingBarFill = document.getElementById('loadingBarFill');
  const loadingText = document.getElementById('loadingText');
  const fontSizeInput = document.getElementById('fontSize');
  const btnPreview = document.getElementById('btnPreview');
  const btnWord = document.getElementById('btnWord');
  const btnExcel = document.getElementById('btnExcel');
  const btnText = document.getElementById('btnText');
  const btnScriptKorean = document.getElementById('btnScriptKorean');
  const btnScriptForeign = document.getElementById('btnScriptForeign');
  const previewAccordion = document.getElementById('previewAccordion');
  const unitsOnlyList = document.getElementById('unitsOnlyList');
  const previewPopup = document.getElementById('previewPopup');
  const previewContent = document.getElementById('previewContent');
  const previewClose = document.getElementById('previewClose');
  const btnRefresh = document.getElementById('btnRefresh');
  const textbookSelect = document.getElementById('textbookSelect');

  let extractedData = [];
  let allUnitsData = [];       // 단원별 전체 라인 (외국인/폴백)
  let allUnitsDataKorean = []; // 단원별 문항 상세 (한국인 성우: 문항번호/지시문/지문/선택지)
  let allPageLines = null;
  let pagesData = [];
  let previewPagesKorean = [];
  let previewPagesForeign = [];
  let pageToUnitLabel = {}; // 목차에서 파싱: 페이지 번호 -> '## 영어듣기 모의고사'
  let scriptType = 'korean';
  let textbookId = 'suneungmanman';
  let uploadedFileName = '';
  /** 미리보기 기반 다운로드용: { fileName, blocks: [{ unitLabel, contentLines }] } */
  let previewExportData = null;

  const pagePreviewPanel = document.getElementById('pagePreviewPanel');
  const pagePreviewList = document.getElementById('pagePreviewList');
  const previewPageList = document.getElementById('previewPageList');
  const previewPageSection = document.getElementById('previewPageSection');
  const fontDownloadArea = document.getElementById('fontDownloadArea');
  const previewArea = document.getElementById('previewArea');
  const appVersionEl = document.getElementById('appVersion');

  function setFontAndPreviewAreasEnabled(enabled) {
    if (fontDownloadArea) fontDownloadArea.classList.toggle('area-disabled', !enabled);
    if (previewArea) previewArea.classList.toggle('area-disabled', !enabled);
    if (fontSizeInput) fontSizeInput.disabled = !enabled;
    if (btnPreview) btnPreview.disabled = !enabled;
    if (btnWord) btnWord.disabled = !enabled;
    if (btnExcel) btnExcel.disabled = !enabled;
    if (btnText) btnText.disabled = !enabled;
  }

  setFontAndPreviewAreasEnabled(false);

  // 버전 표시: config.json 우선, 실패 시 app.js 상수 사용 (file:// 또는 경로 이슈 대비)
  var APP_VERSION = '8.6';
  if (appVersionEl) {
    var configUrl = (document.currentScript && document.currentScript.src)
      ? new URL('../config.json', document.currentScript.src).href
      : 'config.json';
    fetch(configUrl)
      .then(function (res) { return res.ok ? res.json() : Promise.reject(); })
      .then(function (data) {
        appVersionEl.textContent = '(v' + (data.version != null ? data.version : APP_VERSION) + ')';
        if (textbookSelect && data.textbooks && Array.isArray(data.textbooks) && data.textbooks.length > 0) {
          textbookSelect.innerHTML = '';
          data.textbooks.forEach(function (tb) {
            var opt = document.createElement('option');
            opt.value = tb.id || '';
            opt.textContent = tb.name || tb.id || '';
            textbookSelect.appendChild(opt);
          });
        }
      })
      .catch(function () {
        appVersionEl.textContent = '(v' + APP_VERSION + ')';
      });
  }

  if (textbookSelect) {
    textbookSelect.addEventListener('change', function () {
      textbookId = textbookSelect.value || 'suneungmanman';
    });
  }

  function updateUploadNames(files) {
    if (!files || !files.length) {
      uploadFilenames.innerHTML = '';
      return;
    }
    const ul = document.createElement('ul');
    for (let i = 0; i < files.length; i++) {
      const li = document.createElement('li');
      li.textContent = files[i].name;
      ul.appendChild(li);
    }
    uploadFilenames.innerHTML = '';
    uploadFilenames.appendChild(ul);
  }

  function setFiles(files) {
    if (!files || !files.length) return;
    if (files !== pdfInput.files) {
      try {
        var dt = new DataTransfer();
        for (var i = 0; i < files.length; i++) dt.items.add(files[i]);
        pdfInput.files = dt.files;
      } catch (err) {
        if (files.length) pdfInput.files = files;
      }
    }
    btnExtract.disabled = false;
    uploadedFileName = (pdfInput.files[0] && pdfInput.files[0].name || '').replace(/\.pdf$/i, '');
    updateUploadNames(pdfInput.files);
  }

  pdfInput.addEventListener('change', function () {
    if (this.files && this.files.length) setFiles(this.files);
  });
  uploadZone.addEventListener('dragover', function (e) {
    e.preventDefault();
    e.stopPropagation();
    uploadZone.classList.add('dragover');
  });
  uploadZone.addEventListener('dragleave', function (e) {
    e.preventDefault();
    e.stopPropagation();
    uploadZone.classList.remove('dragover');
  });
  uploadZone.addEventListener('drop', function (e) {
    e.preventDefault();
    e.stopPropagation();
    uploadZone.classList.remove('dragover');
    var files = e.dataTransfer.files;
    if (!files || !files.length) return;
    var pdfs = [];
    for (var i = 0; i < files.length; i++) {
      if (files[i].type === 'application/pdf') pdfs.push(files[i]);
    }
    if (pdfs.length) setFiles(pdfs);
  });

  function countDistinctUnits(data) {
    if (!data || !data.length) return 0;
    var seen = {};
    data.forEach(function (r) {
      var key = (r.unitNo != null ? r.unitNo : '') + '|' + (r.unitName != null ? r.unitName : '');
      if (!seen[key]) seen[key] = true;
    });
    return Object.keys(seen).length;
  }

  btnScriptKorean.addEventListener('click', function () {
    scriptType = 'korean';
    btnScriptKorean.classList.add('active');
    btnScriptForeign.classList.remove('active');
    if (allPageLines && allPageLines.length) {
      extractedData = parsePagesForKorean(allPageLines);
      renderAccordion();
    }
    if (pagesData && pagesData.length) {
      renderPagePreviewPanel(pagesData, previewPagesKorean, scriptType);
    }
  });
  btnScriptForeign.addEventListener('click', function () {
    scriptType = 'foreign';
    btnScriptForeign.classList.add('active');
    btnScriptKorean.classList.remove('active');
    if (allPageLines && allPageLines.length) {
      extractedData = parsePagesForForeign(allPageLines);
      renderAccordion();
    }
    if (pagesData && pagesData.length) {
      renderPagePreviewPanel(pagesData, previewPagesForeign, scriptType);
    }
  });

  btnExtract.addEventListener('click', runExtraction);
  btnPreview.addEventListener('click', showFontPreview);
  previewClose.addEventListener('click', function () { previewPopup.classList.remove('show'); });
  previewPopup.addEventListener('click', function (e) {
    if (e.target === previewPopup) previewPopup.classList.remove('show');
  });
  btnWord.addEventListener('click', function () { downloadFile('word'); });
  btnExcel.addEventListener('click', function () { downloadFile('excel'); });
  btnText.addEventListener('click', function () { downloadFile('text'); });
  if (btnRefresh) btnRefresh.addEventListener('click', function () { window.location.reload(); });
  var btnPreviewTopFloat = document.getElementById('btnPreviewTopFloat');
  var btnPagePanelTopFloat = document.getElementById('btnPagePanelTopFloat');
  function updateFloatingTopVisibility() {
    if (btnPreviewTopFloat && previewPageSection) {
      var rect = previewPageSection.getBoundingClientRect();
      var previewScrolled = window.scrollY > (previewPageSection.offsetTop || 0) + 80;
      var previewHasContent = !previewPageSection.classList.contains('empty');
      btnPreviewTopFloat.classList.toggle('show', previewHasContent && previewScrolled);
    }
    if (btnPagePanelTopFloat && pagePreviewPanel) {
      var panelScrolled = pagePreviewPanel.scrollTop > 80;
      var panelHasContent = !pagePreviewPanel.classList.contains('empty');
      btnPagePanelTopFloat.classList.toggle('show', panelHasContent && panelScrolled);
    }
  }
  if (btnPreviewTopFloat) {
    var pageTopEl = document.getElementById('pageTop');
    btnPreviewTopFloat.addEventListener('click', function () {
      var target = pageTopEl || previewPageSection;
      if (target) target.scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  }
  if (btnPagePanelTopFloat && pagePreviewPanel) {
    btnPagePanelTopFloat.addEventListener('click', function () {
      pagePreviewPanel.scrollTop = 0;
    });
  }
  window.addEventListener('scroll', updateFloatingTopVisibility, { passive: true });
  if (pagePreviewPanel) pagePreviewPanel.addEventListener('scroll', updateFloatingTopVisibility, { passive: true });

  function setStatus(msg, type) {
    extractStatus.textContent = msg;
    extractStatus.className = 'status ' + (type || 'info');
    extractStatus.style.display = msg ? 'block' : 'none';
  }

  function getPageTextItems(page) {
    return page.getTextContent().then(function (content) {
      const items = content.items.map(function (item) {
        const y = item.transform[5];
        const x = item.transform[4];
        const w = (item.width != null && item.width > 0) ? item.width : 0;
        return { str: item.str, y: y, x: x, width: w };
      });
      items.sort(function (a, b) {
        const lineA = Math.round(a.y / 4);
        const lineB = Math.round(b.y / 4);
        if (lineA !== lineB) return lineB - lineA;
        return a.x - b.x;
      });
      return items;
    });
  }

  /** 추출할 최대 페이지 수. 0이면 전체. (예: 4면 처음 4페이지만 추출) */
  var MAX_EXTRACT_PAGES = 0;
  /** 같은 줄 내 텍스트 간 가로 간격을 보존: gap이 SPACE_GAP_THRESHOLD 이상이면 공백 2칸으로 구분해 PDF 복사본과 유사하게. */
  var SPACE_GAP_THRESHOLD = 18;
  /** centerX를 넘기면 왼쪽 영역(위→아래) 먼저, 이어서 오른쪽 영역(위→아래) 순으로 줄을 만듦. 반환: { text, y }[] (머리글/바닥글 구분용). */
  function itemsToLines(items, centerX) {
    function groupItemsIntoLines(sortedItems) {
      const out = [];
      let lastY = null;
      let line = [];
      for (let i = 0; i < sortedItems.length; i++) {
        const item = sortedItems[i];
        const y = Math.round(item.y / 3);
        if (lastY !== null && Math.abs(y - lastY) > 2) {
          if (line.length) {
            var txt = joinLineWithSpacing(line);
            if (txt) out.push({ text: txt, y: line[0].y });
          }
          line = [];
        }
        line.push(item);
        lastY = y;
      }
      if (line.length) {
        var txt = joinLineWithSpacing(line);
        if (txt) out.push({ text: txt, y: line[0].y });
      }
      return out;
    }
    if (centerX == null || centerX === undefined) {
      return groupItemsIntoLines(items);
    }
    const left = items.filter(function (item) { return (item.x || 0) < centerX; });
    const right = items.filter(function (item) { return (item.x || 0) >= centerX; });
    function sortTopToBottomThenLeftToRight(a, b) {
      const lineA = Math.round(a.y / 4);
      const lineB = Math.round(b.y / 4);
      if (lineA !== lineB) return lineB - lineA;
      return a.x - b.x;
    }
    left.sort(sortTopToBottomThenLeftToRight);
    right.sort(sortTopToBottomThenLeftToRight);
    return groupItemsIntoLines(left).concat(groupItemsIntoLines(right));
  }
  /** 페이지 높이 기준: 상단은 접두어 없음, 하단은 [페이지/교재명]/[페이지/파일명]/[파일명]/[생성일]만 표기. [바닥글][페이지] 문구 삭제. 공통 적용. */
  function addHeaderFooterLabels(lineResults, pageHeight) {
    if (!lineResults || !lineResults.length) return [];
    var HEADER_Y_RATIO = 0.85;
    var FOOTER_Y_RATIO = 0.2;
    var topY = pageHeight * HEADER_Y_RATIO;
    var bottomY = pageHeight * FOOTER_Y_RATIO;
    return lineResults.map(function (r) {
      if (r.y > topY) return (r.text || '');
      if (r.y < bottomY) {
        var t = (r.text || '').trim();
        if (/\d{4}-\d{2}-\d{2}/.test(t) || /오전|오후\s*\d{1,2}:\d{2}/.test(t)) return '[생성일] ' + t;
        if (/\.indd\b/i.test(t) && /^\d+\s/.test(t)) return '[페이지/파일명] ' + t;
        if (/\.indd\b/i.test(t)) return '[파일명] ' + t;
        if (/수능만만|영어듣기\s*모의고사|^\s*\d+\s+수능만만|1회\s+\d/.test(t)) return '[페이지/교재명] ' + t;
        return t;
      }
      return r.text || '';
    });
  }
  function joinLineWithSpacing(lineItems) {
    if (!lineItems.length) return '';
    lineItems.sort(function (a, b) { return a.x - b.x; });
    let out = lineItems[0].str || '';
    for (let k = 1; k < lineItems.length; k++) {
      const prevEnd = (lineItems[k - 1].x || 0) + (lineItems[k - 1].width || 0);
      const gap = (lineItems[k].x || 0) - prevEnd;
      out += (gap >= SPACE_GAP_THRESHOLD ? '  ' : ' ') + (lineItems[k].str || '');
    }
    return out.trim();
  }

  function renderPageToThumbnailDataUrl(page, maxWidth) {
    var viewport = page.getViewport({ scale: 1 });
    var scale = maxWidth / viewport.width;
    var scaledViewport = page.getViewport({ scale: scale });
    var canvas = document.createElement('canvas');
    canvas.width = scaledViewport.width;
    canvas.height = scaledViewport.height;
    var ctx = canvas.getContext('2d');
    return page.render({ canvasContext: ctx, viewport: scaledViewport }).promise.then(function () {
      return canvas.toDataURL('image/png');
    });
  }

  /** 페이지별 행(썸네일|데이터)을 동일한 DOM으로 채움. showCopyButton이 true면 미리보기에 [복사] 버튼 추가. */
  function fillPageListWithRows(container, pages, showCopyButton) {
    if (!container) return;
    container.innerHTML = '';
    if (!pages || pages.length === 0) return;
    pages.forEach(function (p) {
      var row = document.createElement('div');
      row.className = 'page-row';
      var thumb = document.createElement('div');
      thumb.className = 'page-row-thumb';
      var img = document.createElement('img');
      img.src = p.thumbnailDataUrl || '';
      img.alt = 'p.' + (p.pageNum);
      var numSpan = document.createElement('div');
      numSpan.className = 'page-num';
      numSpan.textContent = (p.pageNum != null && String(p.pageNum).indexOf('~') !== -1) ? ('P.' + p.pageNum) : ('p.' + p.pageNum);
      thumb.appendChild(img);
      thumb.appendChild(numSpan);
      var data = document.createElement('div');
      data.className = 'page-row-data';
      var pre = document.createElement('pre');
      pre.textContent = (p.lines || []).join('\n');
      data.appendChild(pre);
      if (showCopyButton) {
        var copyBtn = document.createElement('button');
        copyBtn.type = 'button';
        copyBtn.className = 'page-row-copy-btn';
        copyBtn.textContent = '복사';
        copyBtn.addEventListener('click', function () {
          var text = (p.lines || []).join('\n');
          if (typeof navigator !== 'undefined' && navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(text).then(function () {
              copyBtn.textContent = '복사됨';
              setTimeout(function () { copyBtn.textContent = '복사'; }, 1500);
            }).catch(function () { fallbackCopy(text, copyBtn); });
          } else {
            fallbackCopy(text, copyBtn);
          }
        });
        data.appendChild(copyBtn);
      }
      row.appendChild(thumb);
      row.appendChild(data);
      container.appendChild(row);
    });
  }

  /** 미리보기 전용: 상단 1~3줄 단원명 패턴을 한 줄로 정규화(01 영어듣기 모의고사, 01 DICTATION, 02 Word & Expressions, 기출 01 DICTATION 등). */
  function normalizeUnitHeaderInPreview(lines) {
    if (!lines || lines.length === 0) return lines;
    var out = lines.slice();
    var l0 = (out[0] || '').trim();
    var l1 = (out[1] || '').trim();
    var l2 = (out[2] || '').trim();
    var numTwo = function (s) { var n = (s || '').replace(/\s/g, ''); return (n.length === 1 ? '0' + n : n); };
    // 3줄: 숫자만 / "영어듣기 ... 모의고사" / 숫자 → "0N 영어듣기 모의고사"
    if (/^\d{1,2}\s*$/.test(l0) && /영어듣기\s+모의고사/i.test(l1) && /^\d{1,2}\s*$/.test(l2)) {
      out[0] = numTwo(l2) + ' 영어듣기 모의고사';
      out.splice(1, 2);
      return out;
    }
    // 2줄: "영어듣기 ... 모의고사" / 숫자 → "0N 영어듣기 모의고사"
    if (/영어듣기\s+모의고사/i.test(l0) && /^\d{1,2}\s*$/.test(l1)) {
      out[0] = numTwo(l1) + ' 영어듣기 모의고사';
      out.splice(1, 1);
      return out;
    }
    // 2줄: "기출 영어듣기" / "01" → "01 기출 영어듣기"
    if (/기출\s+영어듣기/i.test(l0) && /^\d{1,2}\s*$/.test(l1)) {
      out[0] = numTwo(l1) + ' 기출 영어듣기';
      out.splice(1, 1);
      return out;
    }
    // 2줄: "기출" / "01   D ICTATION" 등 → "기출 01 DICTATION"
    if (/^기출\s*$/i.test(l0) && /\d{1,2}\s+D\s*I\s*C\s*T\s*A\s*T\s*I\s*O\s*N/i.test(l1)) {
      var m = l1.match(/^(\d{1,2})\s+/);
      out[0] = '기출 ' + (m ? numTwo(m[1]) : '01') + ' DICTATION';
      out.splice(1, 1);
      return out;
    }
    if (/^기출\s*$/i.test(l0) && /Word\s*&\s*Expressions/i.test(l1)) {
      var m = l1.match(/^(\d{1,2})\s+/);
      out[0] = '기출 ' + (m ? numTwo(m[1]) : '01') + ' Word & Expressions';
      out.splice(1, 1);
      return out;
    }
    // 1줄: "01   D ICTATION" / "01  D ICTATION" → "01 DICTATION"
    if (/^\d{1,2}\s+D\s*I\s*C\s*T\s*A\s*T\s*I\s*O\s*N\s*$/i.test(l0) || /^\d{1,2}\s+DICTATION\s*$/i.test(l0)) {
      var n = l0.match(/^(\d{1,2})/);
      out[0] = (n ? numTwo(n[1]) : '01') + ' DICTATION';
      return out;
    }
    if (/^\d{1,2}\s+Word\s*&\s*Expressions\s*$/i.test(l0) || /^\d{1,2}\s+W\s*o\s*r\s*d\s*&\s*E\s*x\s*p\s*r\s*e\s*s\s*s\s*i\s*o\s*n\s*s\s*$/i.test(l0)) {
      var n = l0.match(/^(\d{1,2})/);
      out[0] = (n ? numTwo(n[1]) : '01') + ' Word & Expressions';
      return out;
    }
    if (/^\d{1,2}\s+영어듣기\s+모의고사\s*$/i.test(l0)) {
      var n = l0.match(/^(\d{1,2})/);
      out[0] = (n ? numTwo(n[1]) : '01') + ' 영어듣기 모의고사';
      return out;
    }
    // 2줄: "01" / "D ICTATION" 또는 "Word & Expressions" (공백 다양)
    if (/^\d{1,2}\s*$/.test(l0) && /D\s*I\s*C\s*T\s*A\s*T\s*I\s*O\s*N/i.test(l1)) {
      out[0] = numTwo(l0) + ' DICTATION';
      out.splice(1, 1);
      return out;
    }
    if (/^\d{1,2}\s*$/.test(l0) && /Word\s*&\s*Expressions/i.test(l1)) {
      out[0] = numTwo(l0) + ' Word & Expressions';
      out.splice(1, 1);
      return out;
    }
    return out;
  }

  /** 미리보기 전용: 한 줄이 단원명인지 판별(01 영어듣기 모의고사, 01 DICTATION, 기출 01 DICTATION 등). */
  function isPreviewUnitNameLine(line) {
    var t = (line || '').trim();
    return /^\d{1,2}\s+영어듣기\s+모의고사\s*$/i.test(t) || /^\d{1,2}\s+DICTATION\s*$/i.test(t) || /^\d{1,2}\s+Word\s*&\s*Expressions\s*$/i.test(t) || /^\d{1,2}\s+기출\s+영어듣기\s*$/i.test(t) || /^기출\s+\d{1,2}\s+DICTATION\s*$/i.test(t) || /^기출\s+\d{1,2}\s+Word\s*&\s*Expressions\s*$/i.test(t);
  }

  /** 미리보기 블록을 다운로드용 구조로 변환: [{ unitLabel, contentLines }]. fillPreviewPagesNoThumb과 동일한 전처리 적용. */
  function getPreviewExportBlocks(pagesOrBlocks, scriptType) {
    if (!pagesOrBlocks || !pagesOrBlocks.length) return [];
    var out = [];
    pagesOrBlocks.forEach(function (p) {
      var filteredLines;
      if (p.pageRange != null) {
        filteredLines = p.lines || [];
      } else {
        filteredLines = stripPreviewMetadataLines(p.lines || []);
        filteredLines = normalizeUnitHeaderInPreview(filteredLines);
      }
      var firstLine = (filteredLines[0] || '').trim();
      var useUnitInTitle = filteredLines.length > 0 && isPreviewUnitNameLine(firstLine);
      var unitLabel = useUnitInTitle ? firstLine : '';
      var contentLines = useUnitInTitle ? filteredLines.slice(1) : filteredLines;
      contentLines = stripTrailingPageNumber(contentLines || []);
      contentLines = stripPreviewInstructionLines(contentLines || []);
      contentLines = mergePreviewInstructionToSingleLine(contentLines || []);
      if (scriptType === 'foreign') {
        var proc = processForeignPreviewContent((contentLines || []).map(function (ln) { return normalizeKoreanPreviewLine(ln); }));
        contentLines = proc.contentLines || [];
      } else {
        contentLines = (contentLines || []).map(function (ln) { return normalizeKoreanPreviewLine(ln); });
        if (scriptType === 'korean') contentLines = reduceKoreanPreviewToQuestionInstructionOnly(contentLines);
      }
      out.push({ unitLabel: unitLabel, contentLines: contentLines || [] });
    });
    return out;
  }

  /** 미리보기 전용: 썸네일 없이 블록별 표시. 타이틀은 '단원명 (p.페이지번호)', 단원명이 타이틀에 쓰이면 본문에서는 비노출. scriptType === 'korean'이면 문항번호+지시문 1줄만, foreign이면 W:/M: 색상·문항번호 볼드 적용. */
  function fillPreviewPagesNoThumb(container, pagesOrBlocks, scriptType) {
    if (!container) return;
    container.innerHTML = '';
    if (!pagesOrBlocks || pagesOrBlocks.length === 0) return;
    pagesOrBlocks.forEach(function (p) {
      var filteredLines;
      if (p.pageRange != null) {
        filteredLines = p.lines || [];
      } else {
        filteredLines = stripPreviewMetadataLines(p.lines || []);
        filteredLines = normalizeUnitHeaderInPreview(filteredLines);
      }
      var range = p.pageRange != null ? p.pageRange : p.pageNum;
      var rangeStr = (range != null && String(range).indexOf('~') !== -1) ? ('p.' + range) : ('p.' + (range != null ? range : ''));
      var firstLine = (filteredLines[0] || '').trim();
      var useUnitInTitle = filteredLines.length > 0 && isPreviewUnitNameLine(firstLine);
      var titleDisplay = useUnitInTitle ? (firstLine + ' (' + rangeStr + ')') : rangeStr;
      var contentLines = useUnitInTitle ? filteredLines.slice(1) : filteredLines;
      contentLines = stripTrailingPageNumber(contentLines || []);
      contentLines = stripPreviewInstructionLines(contentLines || []);
      contentLines = mergePreviewInstructionToSingleLine(contentLines || []);
      contentLines = (contentLines || []).map(function (ln) { return normalizeKoreanPreviewLine(ln); });
      var text;
      var foreignProc = null;
      if (scriptType === 'foreign') {
        foreignProc = processForeignPreviewContent(contentLines);
        contentLines = foreignProc.contentLines;
        text = (contentLines || []).join('\n');
      } else {
        if (scriptType === 'korean') contentLines = reduceKoreanPreviewToQuestionInstructionOnly(contentLines);
        text = (contentLines || []).join('\n');
      }
      var row = document.createElement('div');
      row.className = 'page-row page-row--no-thumb';
      var data = document.createElement('div');
      data.className = 'page-row-data';
      var titleRow = document.createElement('div');
      titleRow.className = 'page-row-title';
      var titleText = document.createElement('span');
      titleText.className = 'page-row-title-text';
      titleText.textContent = titleDisplay;
      var copyBtn = document.createElement('button');
      copyBtn.type = 'button';
      copyBtn.className = 'page-row-copy-btn';
      copyBtn.textContent = '복사';
      copyBtn.addEventListener('click', function () {
        if (typeof navigator !== 'undefined' && navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(text).then(function () {
            copyBtn.textContent = '복사됨';
            setTimeout(function () { copyBtn.textContent = '복사'; }, 1500);
          }).catch(function () { fallbackCopy(text, copyBtn); });
        } else {
          fallbackCopy(text, copyBtn);
        }
      });
      titleRow.appendChild(titleText);
      titleRow.appendChild(copyBtn);
      data.appendChild(titleRow);
      if (scriptType === 'foreign' && foreignProc) {
        var contentDiv = document.createElement('div');
        contentDiv.className = 'page-row-foreign-content';
        contentDiv.style.whiteSpace = 'pre-wrap';
        contentDiv.style.wordBreak = 'break-word';
        contentDiv.style.lineHeight = '1.45';
        contentDiv.style.margin = '0';
        contentDiv.style.padding = '0.5rem';
        contentDiv.style.maxHeight = '360px';
        contentDiv.style.overflowY = 'auto';
        contentDiv.style.background = '#fff';
        contentDiv.style.borderRadius = '6px';
        contentDiv.style.border = '1px solid ' + (getComputedStyle ? getComputedStyle(document.body).getPropertyValue('--border') || '#e2e8f0' : '#e2e8f0');
        contentDiv.style.fontSize = '0.75rem';
        contentDiv.style.fontFamily = 'var(--font), NanumSquareNeo, sans-serif';
        contentDiv.innerHTML = foreignProc.displayHtml;
        data.appendChild(contentDiv);
      } else {
        var pre = document.createElement('pre');
        pre.textContent = text;
        data.appendChild(pre);
      }
      row.appendChild(data);
      container.appendChild(row);
    });
  }

  /** 한국인 성우 미리보기: 썸네일 없음, 단원명 + [복사] 상단, 그 아래 추출 데이터(단원명 제외) */
  function fillKoreanPreviewRows(container, units) {
    if (!container) return;
    container.innerHTML = '';
    if (!units || units.length === 0) return;
    units.forEach(function (u) {
      var allLines = u.lines || [];
      var dataLines = allLines.slice();
      if (u.unitLabel && dataLines.length && (dataLines[0] || '').trim() === (u.unitLabel || '').trim()) {
        dataLines = dataLines.slice(1);
      }
      var dataText = dataLines.join('\n');
      var row = document.createElement('div');
      row.className = 'page-row page-row--no-thumb';
      var data = document.createElement('div');
      data.className = 'page-row-data';
      var titleRow = document.createElement('div');
      titleRow.className = 'page-row-title';
      var titleText = document.createElement('span');
      titleText.className = 'page-row-title-text';
      titleText.textContent = u.unitLabel || '';
      var copyBtn = document.createElement('button');
      copyBtn.type = 'button';
      copyBtn.className = 'page-row-copy-btn';
      copyBtn.textContent = '복사';
      copyBtn.addEventListener('click', function () {
        var text = dataText;
        if (typeof navigator !== 'undefined' && navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(text).then(function () {
            copyBtn.textContent = '복사됨';
            setTimeout(function () { copyBtn.textContent = '복사'; }, 1500);
          }).catch(function () { fallbackCopy(text, copyBtn); });
        } else {
          fallbackCopy(text, copyBtn);
        }
      });
      titleRow.appendChild(titleText);
      titleRow.appendChild(copyBtn);
      data.appendChild(titleRow);
      var pre = document.createElement('pre');
      pre.textContent = dataText;
      data.appendChild(pre);
      row.appendChild(data);
      container.appendChild(row);
    });
  }

  function fallbackCopy(text, btn) {
    var ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    try {
      document.execCommand('copy');
      btn.textContent = '복사됨';
      setTimeout(function () { btn.textContent = '복사'; }, 1500);
    } catch (e) {}
    document.body.removeChild(ta);
  }

  /** 페이지별 추출(오른쪽)과 미리보기(왼쪽) 동일하게 페이지 단위로 표시. 오른쪽은 썸네일+데이터, 왼쪽은 썸네일 없이 동일 페이지 순서·내용. */
  function renderPagePreviewPanel(allPages, previewPages, scriptType) {
    if (!pagePreviewPanel || !pagePreviewList) return;
    var hasAll = allPages && allPages.length > 0;

    if (!hasAll) {
      pagePreviewPanel.classList.add('empty');
      pagePreviewList.innerHTML = '';
      if (typeof updateFloatingTopVisibility === 'function') updateFloatingTopVisibility();
    } else {
      pagePreviewPanel.classList.remove('empty');
      fillPageListWithRows(pagePreviewList, allPages);
    }

    if (!previewPageSection) return;
    if (!hasAll) {
      previewPageSection.classList.add('empty');
      if (previewPageList) previewPageList.innerHTML = '';
      if (typeof updateFloatingTopVisibility === 'function') updateFloatingTopVisibility();
    } else {
      previewPageSection.classList.remove('empty');
      // 미리보기: 구성과 특징·목차 등 비노출, 단원명 없는 페이지는 앞 단원 마지막에 병합
      var pagesForPreview = allPages.filter(function (p) { return !isExcludedFromPreview(p.lines || []); });
      var mergedBlocks = mergePreviewPagesByUnit(pagesForPreview);
      // 스크립트 유형별 필터: 한국인 성우=지시문 있는 단원만, 외국인 성우=W:/M: 대화 지문 있는 단원만
      if (scriptType === 'korean') {
        mergedBlocks = mergedBlocks.filter(function (b) { return hasInstructionPhrase(b.lines || []); });
        mergedBlocks = mergedBlocks.map(function (b) {
          return { pageNum: b.pageNum, pageRange: b.pageRange, lines: (b.lines || []).map(normalizeKoreanPreviewLine) };
        });
      } else if (scriptType === 'foreign') {
        mergedBlocks = mergedBlocks.filter(function (b) { return hasDialogueWM(b.lines || []); });
      }
      var previewUnitCount = mergedBlocks.length;
      if (scriptType === 'korean') {
        setStatus('한국인 성우 스크립트 기준 총 ' + previewUnitCount + ' 단원.', 'success');
      } else if (scriptType === 'foreign') {
        setStatus('외국인 성우 스크립트 기준 총 ' + previewUnitCount + ' 단원.', 'success');
      }
      previewExportData = { fileName: uploadedFileName || '', blocks: getPreviewExportBlocks(mergedBlocks, scriptType) };
      fillPreviewPagesNoThumb(previewPageList, mergedBlocks, scriptType);
      if (typeof updateFloatingTopVisibility === 'function') updateFloatingTopVisibility();
    }
  }

  function isCoverOrToc(lines, pageIndex) {
    var text = lines.join(' ');
    if (lines.length < 3) return true;
    if (pageIndex != null && pageIndex >= 20) return false;
    if (/목차|차례|표지|특징|구성/.test(text) && lines.length < 12) return true;
    return false;
  }

  function isExcludedPage(lines, pageIndex) {
    var text = lines.join(' ');
    if (/표지/.test(text) && lines.length < 12) return true;
    if (/구성과\s*특징|구성\s*및\s*특징/.test(text) && lines.length < 15) return true;
    if (/목차|차례/.test(text) && lines.length < 12) return true;
    if (/간지/.test(text) && lines.length < 10) return true;
    return false;
  }

  /** 미리보기 전용: [페이지/파일명], [생성일], [페이지/교재명] 라인 제거 */
  function stripPreviewMetadataLines(lines) {
    if (!lines || !lines.length) return lines;
    var skipRe = /^\s*\[(?:페이지\/파일명|생성일|페이지\/교재명)\]\s*/;
    return lines.filter(function (line) { return !skipRe.test((line || '').trim()); });
  }

  /** 미리보기 전용: 마지막에 있는 페이지 번호 제거 (단독 숫자 줄, 마지막 줄 끝 숫자) */
  function stripTrailingPageNumber(lines) {
    if (!lines || !lines.length) return lines;
    var out = lines.slice();
    while (out.length > 0 && /^\s*\d{1,3}\s*$/.test((out[out.length - 1] || '').trim())) {
      out.pop();
    }
    if (out.length > 0) {
      var last = (out[out.length - 1] || '').replace(/\s+\d{1,3}\s*$/, '');
      if (last !== out[out.length - 1]) out[out.length - 1] = last.trim();
    }
    return out;
  }

  /** 미리보기 전용: 지시문 2줄 이상을 1줄로 합치고, 끝의 'N 회' 제거, 공백 정리 */
  function mergePreviewInstructionToSingleLine(lines) {
    if (!lines || !lines.length) return lines;
    var result = [];
    var i = 0;
    while (i < lines.length) {
      var line = (lines[i] || '').trim();
      if (!line) { i++; continue; }
      if (/^\d{1,2}\s*$/.test(line)) {
        result.push(line);
        i++;
        continue;
      }
      var merged = line;
      var j = i + 1;
      while (j < lines.length) {
        var nextLine = (lines[j] || '').trim();
        if (/^\d{1,2}\s*$/.test(nextLine) || /^\d{1,2}\s*[~\-–−]\s*\d{1,2}\s*$/.test(nextLine)) break;
        if (!nextLine) { j++; continue; }
        merged = (merged + ' ' + nextLine).replace(/\s{2,}/g, ' ').trim();
        merged = merged.replace(/\s*\d+\s*회\s*$/g, '').replace(/\s+\d+\s*회\s*/g, ' ').trim();
        j++;
        if (isRealInstructionEnd(merged)) break;
      }
      merged = merged.replace(/\s*\d+\s*회\s*$/g, '').replace(/\s+\d+\s*회\s+/g, ' ').trim();
      result.push(merged);
      i = j;
    }
    return result;
  }

  /** 한국인 성우 스크립트 미리보기: '고르시 오.' '고르 시오.' '고 르시오.' → '고르시오.' 로 수정, 띄어쓰기 2칸 이상→1칸 */
  function normalizeKoreanPreviewLine(line) {
    if (line == null || line === '') return line;
    var t = String(line).replace(/\s{2,}/g, ' ').trim();
    t = t.replace(/고\s*르\s*시\s*오\s*\.?/g, '고르시오.');
    return t;
  }

  /** 외국인 성우: 문항번호 라인인지 (01, 02, 16~17, 16-17 등, en-dash 포함) */
  function isForeignQuestionNumberLine(line) {
    if (!line || typeof line !== 'string') return false;
    var t = line.trim();
    if (/^\d{1,2}\s*$/.test(t)) return true;
    if (/^\d{1,2}\s*[~\-–−]\s*\d{1,2}\s*$/.test(t)) return true;
    return false;
  }

  /** 외국인 성우: 02번부터 또는 16~17/16-17 이면 앞줄 띄기 필요 */
  function needsBlankBeforeQuestionNum(line) {
    if (!line || typeof line !== 'string') return false;
    var t = line.trim();
    if (/^01\s*$/.test(t)) return false;
    if (/^\d{1,2}\s*$/.test(t)) return true;
    if (/^\d{1,2}\s*[~\-–−]\s*\d{1,2}\s*$/.test(t)) return true;
    return false;
  }

  /** 외국인 성우: 라인에서 문항번호+대화 분리 (예: "16~17 M: Welcome..." → ["16~17", "M: Welcome..."]) */
  function splitQuestionNumAndDialogue(line) {
    var t = (line || '').trim();
    if (!t) return null;
    var m = t.match(/^(\d{1,2}\s*[~\-–−]\s*\d{1,2}|\d{1,2})\s+([WM]:\s*.+)$/);
    if (m) return [m[1].trim(), m[2].trim()];
    return [t];
  }

  /** 외국인 성우: 라인 끝의 16-17, 16~17 분리 (예: "Frances: (...) 16-17" → ["Frances: (...)", "16-17"], 공백 없이 붙어도 분리) */
  function splitTrailingQuestionNum(line) {
    var t = (line || '').trim();
    if (!t) return [t];
    var m = t.match(/^(.+?)\s*(16\s*[~\-–−]\s*17)\s*$/);
    if (m) return [m[1].trim(), m[2].trim()];
    return [t];
  }

  /** 외국인 성우: 발화자 태그로 시작하는지 (M:, W:, 영어이름:) */
  function startsWithSpeakerTag(text) {
    if (!text || typeof text !== 'string') return null;
    var t = text.trim();
    var m = t.match(/^(M|W)\s*:\s*(.*)$/);
    if (m) return m[1].toLowerCase();
    var n = t.match(/^([A-Z][a-z]+)\s*:\s*(.*)$/);
    if (n && n[1] !== 'M' && n[1] !== 'W') return 'other';
    return null;
  }

  /** 외국인 성우: M:, W:, 영어이름: 기준으로 텍스트를 블록 배열로 분리 [{ speaker, text }] */
  function splitBySpeakerBlocks(text) {
    if (!text || typeof text !== 'string') return [];
    var t = text.replace(/\s{2,}/g, ' ').trim();
    if (!t) return [];
    var re = /(?=(?:M|W)\s*:\s*|\b(?:[A-Z][a-z]+)\s*:\s*)/g;
    var parts = t.split(re).map(function (p) { return p.trim(); }).filter(Boolean);
    var result = [];
    for (var i = 0; i < parts.length; i++) {
      var p = parts[i];
      var sp = startsWithSpeakerTag(p);
      if (sp) result.push({ speaker: sp, text: p });
      else if (result.length) result[result.length - 1].text += (result[result.length - 1].text ? ' ' : '') + p;
      else result.push({ speaker: 'other', text: p });
    }
    return result;
  }

  /** 외국인 성우 미리보기: M:/W:/영어이름: 나오기 전까지 같은 색상 블록, 16-17/16~17 앞 한줄띄기 */
  function processForeignPreviewContent(lines) {
    var contentLines = [];
    var displayParts = [];
    var COLOR_M = '#1565c0';
    var COLOR_W = '#c2185b';
    var COLOR_OTHER = '#2e7d32';
    var tokens = [];
    for (var i = 0; i < lines.length; i++) {
      var line = (lines[i] || '').trim();
      if (!line) continue;
      var trailingParts = splitTrailingQuestionNum(line);
      var rest = trailingParts[0];
      var trailing = trailingParts[1];
      if (trailing && rest !== line) {
        if (isForeignQuestionNumberLine(rest)) tokens.push({ type: 'qnum', text: rest });
        else tokens.push({ type: 'text', text: rest });
        tokens.push({ type: 'qnum', text: trailing });
        continue;
      }
      var segs = splitQuestionNumAndDialogue(line);
      for (var s = 0; s < segs.length; s++) {
        var seg = segs[s];
        if (isForeignQuestionNumberLine(seg)) {
          tokens.push({ type: 'qnum', text: seg });
        } else {
          tokens.push({ type: 'text', text: seg });
        }
      }
    }
    var blockBuf = '';
    var blockSpeaker = null;
    for (var t = 0; t < tokens.length; t++) {
      var tok = tokens[t];
      if (tok.type === 'qnum') {
        if (blockBuf && blockSpeaker) {
          contentLines.push(blockBuf.trim());
          displayParts.push({ type: blockSpeaker, text: blockBuf.trim() });
          blockBuf = '';
          blockSpeaker = null;
        }
        if (needsBlankBeforeQuestionNum(tok.text) && contentLines.length > 0) {
          contentLines.push('');
          displayParts.push('');
        }
        contentLines.push(tok.text);
        displayParts.push({ type: 'qnum', text: tok.text });
      } else {
        if (startsWithSpeakerTag(tok.text)) {
          var blocks = splitBySpeakerBlocks(tok.text);
          for (var b = 0; b < blocks.length; b++) {
            var blk = blocks[b];
            if (blockBuf && blockSpeaker && blockSpeaker !== blk.speaker) {
              contentLines.push(blockBuf.trim());
              displayParts.push({ type: blockSpeaker, text: blockBuf.trim() });
              blockBuf = '';
            }
            if (blockSpeaker === blk.speaker && blockBuf) {
              blockBuf += ' ' + blk.text.replace(/\s{2,}/g, ' ').trim();
            } else {
              if (blockBuf && blockSpeaker) {
                contentLines.push(blockBuf.trim());
                displayParts.push({ type: blockSpeaker, text: blockBuf.trim() });
              }
              blockBuf = blk.text;
              blockSpeaker = blk.speaker;
            }
          }
        } else {
          if (blockBuf) blockBuf += ' ' + tok.text.replace(/\s{2,}/g, ' ').trim();
          else { blockBuf = tok.text; if (!blockSpeaker) blockSpeaker = 'other'; }
        }
      }
    }
    if (blockBuf && blockSpeaker) {
      contentLines.push(blockBuf.trim());
      displayParts.push({ type: blockSpeaker, text: blockBuf.trim() });
    }
    var collapsedLines = [];
    var collapsedParts = [];
    for (var k = 0; k < contentLines.length; k++) {
      if (contentLines[k] === '') {
        if (collapsedLines[collapsedLines.length - 1] === '') continue;
      }
      collapsedLines.push(contentLines[k]);
      collapsedParts.push(displayParts[k]);
    }
    contentLines = collapsedLines;
    displayParts = collapsedParts;
    /* 16~17/16-17 앞에 빈 줄 보장 (이중 안전장치) */
    var finalLines = [];
    var finalParts = [];
    for (var idx = 0; idx < contentLines.length; idx++) {
      var ln = contentLines[idx];
      if (/^\d{1,2}\s*[~\-–−]\s*\d{1,2}\s*$/.test((ln || '').trim())) {
        if (finalLines.length > 0 && finalLines[finalLines.length - 1] !== '') {
          finalLines.push('');
          finalParts.push('');
        }
      }
      finalLines.push(ln);
      finalParts.push(displayParts[idx]);
    }
    contentLines = finalLines;
    displayParts = finalParts;
    var displayHtml = displayParts.map(function (p) {
      if (p === '') return '<br>';
      if (typeof p === 'object') {
        if (p.type === 'qnum') return '<span class="foreign-qnum" style="font-weight:bold">' + escapeHtml(p.text) + '</span>';
        if (p.type === 'm') return '<span class="foreign-m" style="color:' + COLOR_M + '">' + escapeHtml(p.text) + '</span>';
        if (p.type === 'w') return '<span class="foreign-w" style="color:' + COLOR_W + '">' + escapeHtml(p.text) + '</span>';
        if (p.type === 'other') return '<span class="foreign-other" style="color:' + COLOR_OTHER + '">' + escapeHtml(p.text) + '</span>';
        return '<span>' + escapeHtml(p.text) + '</span>';
      }
      return '<span>' + escapeHtml(p) + '</span>';
    }).join('<br>');
    return { contentLines: contentLines, displayHtml: displayHtml };
  }

  /** 지시문 끝 위치: 마침표(.)·물음표(?) 중 첫 번째. 단 '알파벳+마침표'(Mrs., Mr., Dr. 등)는 제외 */
  function findInstructionEndIndex(str) {
    if (!str || typeof str !== 'string') return -1;
    for (var i = 0; i < str.length; i++) {
      if (str[i] === '?' || str[i] === '\uFF1F') return i;
      if (str[i] === '.') {
        var j = i - 1;
        while (j >= 0 && /[A-Za-z]/.test(str[j])) j--;
        if (i - 1 - j > 0) continue;
        return i;
      }
    }
    return -1;
  }

  /** 문자열이 지시문 끝(실제 마침표·물음표)으로 끝나는지. '알파벳+마침표'는 제외 */
  function isRealInstructionEnd(str) {
    if (!str || typeof str !== 'string') return false;
    var t = str.trim();
    if (/[?？]$/.test(t)) return true;
    if (!/\.$/.test(t)) return false;
    var i = t.length - 1;
    var j = i - 1;
    while (j >= 0 && /[A-Za-z]/.test(t[j])) j--;
    return (i - 1 - j) === 0;
  }

  /** 한국인 성우 미리보기: 문항번호 + 지시문 1줄만 남기고 나머지(W:, M:, 선지 등) 비노출. 지시문=대화를 듣고/다음을 듣고/고르시오/않은 것은/아닌 것은/적절한 것 포함, 마침표·물음표까지(알파벳+마침표 제외) */
  function reduceKoreanPreviewToQuestionInstructionOnly(lines) {
    if (!lines || !lines.length) return [];
    var instructionKeyword = /다음을\s*듣고|대화를\s*듣고|고르시오|않은\s*것은|아닌\s*것은|적절한\s*것/;
    var result = [];
    var i = 0;
    while (i < lines.length) {
      var line = (lines[i] || '').trim();
      if (!line) { i++; continue; }
      if (/^\d{1,2}\s*$/.test(line)) {
        var num = line.replace(/\s/g, '');
        i++;
        var instrParts = [];
        while (i < lines.length) {
          var l = (lines[i] || '').trim();
          if (!l) { i++; continue; }
          if (/^\d{1,2}\s*$/.test(l)) break;
          if (instructionKeyword.test(l)) {
            instrParts.push(l);
            i++;
            var combined = instrParts.join(' ').replace(/\s{2,}/g, ' ').trim();
            if (isRealInstructionEnd(combined)) break;
          } else break;
        }
        if (instrParts.length) {
          var combined = instrParts.join(' ').replace(/\s{2,}/g, ' ').trim();
          var endIdx = findInstructionEndIndex(combined);
          if (endIdx >= 0) combined = combined.substring(0, endIdx + 1);
          result.push(num + ' ' + combined);
        }
        continue;
      }
      var m = line.match(/^(\d{1,2})\s+(.+)$/);
      if (m && instructionKeyword.test(m[2])) {
        var rest = m[2].trim();
        var endIdx = findInstructionEndIndex(rest);
        if (endIdx >= 0) rest = rest.substring(0, endIdx + 1);
        result.push(m[1] + ' ' + rest);
      }
      i++;
    }
    return result;
  }

  /** 미리보기 전용: 안내 문구·정답 및 해설 라인 제거 */
  function stripPreviewInstructionLines(lines) {
    if (!lines || !lines.length) return lines;
    var skipPatterns = [
      /1번부터\s*17번까지는\s*듣고\s*답하는\s*문제입니다/,
      /1번부터\s*15번까지는\s*한\s*번만\s*들려주고/,
      /16번부터\s*17번까지는\s*두\s*번\s*들려줍니다/,
      /방송을\s*잘\s*듣고\s*답을\s*하기\s*바랍니다/,
      /녹음을\s*다시\s*한\s*번\s*듣고,\s*빈칸에\s*알맞은\s*말을\s*쓰시오/,
      /정답\s*및\s*해설\s*(p\.\s*\d+)?/,
      /\[\s*16\s*-\s*17\s*\]\s*다음을\s*듣고,\s*물음에\s*답하시오/,
      /\[\s*16\s*~\s*17\s*\]\s*다음을\s*듣고,\s*물음에\s*답하시오/
    ];
    return lines.filter(function (line) {
      var t = (line || '').trim();
      if (!t) return true;
      for (var i = 0; i < skipPatterns.length; i++) {
        if (skipPatterns[i].test(t)) return false;
      }
      return true;
    });
  }

  /** 미리보기 전용: 첫 1~2줄에서 단원명 감지 후 정규화 라벨 반환. 없으면 null */
  function detectUnitFromPageLines(lines) {
    if (!lines || lines.length === 0) return null;
    var l0 = (lines[0] || '').trim();
    var l1 = (lines[1] || '').trim();
    var m = l0.match(/영어듣기\s*모의고사\s*$/i);
    if (m && /^\d{1,2}\s*$/.test(l1)) {
      var num = l1.replace(/\s/g, '');
      return (num.length === 1 ? '0' + num : num) + ' 영어듣기 모의고사';
    }
    if (/^\d{1,2}\s+영어듣기\s*모의고사\s*$/i.test(l0)) {
      var n = l0.match(/^(\d{1,2})/)[1];
      return (n.length === 1 ? '0' + n : n) + ' 영어듣기 모의고사';
    }
    var dict = l0.match(/^(\d{1,2})\s+D\s*I\s*C\s*T\s*A\s*T\s*I\s*O\s*N\s*$/i) || l0.match(/^(\d{1,2})\s+DICTATION\s*$/i);
    if (dict) return (dict[1].length === 1 ? '0' + dict[1] : dict[1]) + ' DICTATION';
    var word = l0.match(/^(\d{1,2})\s+Word\s*&\s*Expressions\s*$/i) || l0.match(/^(\d{1,2})\s+W\s*o\s*r\s*d\s*&\s*E\s*x\s*p\s*r\s*e\s*s\s*s\s*i\s*o\s*n\s*s\s*$/i);
    if (word) return (word[1].length === 1 ? '0' + word[1] : word[1]) + ' Word & Expressions';
    if (/^\d{1,2}\s+Word\s+&\s+Expressions/i.test(l0) || /Word\s*&\s*Expressions/i.test(l0)) {
      var w = l0.match(/^(\d{1,2})/);
      if (w) return (w[1].length === 1 ? '0' + w[1] : w[1]) + ' Word & Expressions';
    }
    return null;
  }

  /** 미리보기 전용: 병합된 단원 내용 앞부분 단원명 정규화(영어듣기 모의고사+숫자, DICTATION, Word & Expressions) */
  function normalizeUnitNameInMergedLines(lines, unitLabel) {
    if (!lines || !lines.length || !unitLabel) return lines;
    var out = lines.slice();
    var l0 = (out[0] || '').trim();
    var l1 = (out[1] || '').trim();
    if (/영어듣기\s*모의고사\s*$/i.test(l0) && /^\d{1,2}\s*$/.test(l1)) {
      out[0] = unitLabel;
      out.splice(1, 1);
      return out;
    }
    if (/^\d{1,2}\s+영어듣기\s*모의고사\s*$/i.test(l0)) {
      out[0] = unitLabel;
      return out;
    }
    if (/^\d{1,2}\s+D\s*I\s*C\s*T\s*A\s*T\s*I\s*O\s*N\s*$/i.test(l0) || /^\d{1,2}\s+DICTATION\s*$/i.test(l0)) {
      out[0] = unitLabel;
      return out;
    }
    if (/^\d{1,2}\s+Word\s*&\s*Expressions\s*$/i.test(l0) || /Word\s*&\s*Expressions/i.test(l0)) {
      var w = l0.match(/^(\d{1,2})\s*/);
      if (w) { out[0] = unitLabel; return out; }
    }
    return out;
  }

  /** 미리보기 전용: 표지/구성/목차 제외, 메타라인 제거, 단원 기준 병합, 단원명 정규화. 공통 적용. */
  function buildPreviewUnits(pages, pageToUnitLabel) {
    if (!pages || !pages.length) return [];
    var filtered = pages.filter(function (p) { return !isExcludedFromPreview(p.lines || []); });
    if (!filtered.length) return [];
    var units = [];
    var currentUnit = null;
    for (var i = 0; i < filtered.length; i++) {
      var p = filtered[i];
      var lines = stripPreviewMetadataLines(p.lines || []);
      var label = pageToUnitLabel[p.pageNum] || detectUnitFromPageLines(lines);
      if (label) {
        currentUnit = { unitLabel: label, lines: lines.slice() };
        currentUnit.lines = normalizeUnitNameInMergedLines(currentUnit.lines, label);
        units.push(currentUnit);
      } else if (currentUnit) {
        currentUnit.lines.push('');
        currentUnit.lines = currentUnit.lines.concat(lines);
      } else {
        currentUnit = { unitLabel: 'p.' + p.pageNum, lines: lines.slice() };
        units.push(currentUnit);
      }
    }
    return units.filter(function (u) { return u.lines && u.lines.length; });
  }

  /** 미리보기 전용: 단원명이 없는 페이지는 이전 단원 블록 마지막에 병합. 반환: { pageNum, pageRange, lines }[] */
  function mergePreviewPagesByUnit(pages) {
    if (!pages || !pages.length) return [];
    function hasUnitName(lines) {
      var stripped = stripPreviewMetadataLines(lines || []);
      if (detectUnitFromPageLines(stripped)) return true;
      var norm = normalizeUnitHeaderInPreview(stripped.slice());
      var first = (norm[0] || '').trim();
      return /^\d{1,2}\s+영어듣기\s+모의고사/i.test(first) || /^\d{1,2}\s+DICTATION/i.test(first) || /^\d{1,2}\s+Word\s*&\s*Expressions/i.test(first) || /^기출\s+\d{1,2}\s+DICTATION/i.test(first) || /^기출\s+\d{1,2}\s+Word/i.test(first) || /^\d{1,2}\s+기출\s+영어듣기/i.test(first);
    }
    var result = [];
    for (var i = 0; i < pages.length; i++) {
      var p = pages[i];
      var rawLines = p.lines || [];
      var stripped = stripPreviewMetadataLines(rawLines.slice());
      var normalized = normalizeUnitHeaderInPreview(stripped);
      if (hasUnitName(rawLines) || result.length === 0) {
        result.push({ pageNum: p.pageNum, pageRange: String(p.pageNum), lines: normalized });
      } else {
        var last = result[result.length - 1];
        last.lines = last.lines.concat(['']).concat(normalized);
        last.pageRange = last.pageNum + '~' + p.pageNum;
      }
    }
    return result;
  }

  /** 미리보기에서 제외: 표지, 간지, 구성과 특징, 목차, 지은이 등 */
  function isExcludedFromPreview(lines) {
    var text = lines.join(' ');
    if (/\.indd\b/.test(text) && lines.length < 12) return true;
    if (/구성과\s*특징/.test(text)) return true;
    if (/목차/.test(text)) return true;
    if (/지은이/.test(text)) return true;
    // 표지: 만만한 + 수능영어/영어듣기 + 35+5회 등
    if (/만만한/.test(text) && (/수능영어|영어듣기|영어\s*듣기\s*35\s*\+\s*5\s*회|\d+\s*\+\s*\d+\s*회/.test(text) || lines.length < 15)) return true;
    // 간지: 글자 사이 공백 패턴 (수 능 만 만, 영 어 듣 기, 모 의 고 사 등)
    if (/수\s+능\s+만\s+만|영\s+어\s+듣\s+기|모\s+의\s+고\s+사/.test(text)) return true;
    if (lines.length <= 8 && /\b기\b/.test(text) && /\b본\b/.test(text)) return true;
    if (lines.length <= 10 && /0\s*1\s*$|-\s*35\s*회/.test(text) && /영\s*어|모\s*의/.test(text)) return true;
    return false;
  }

  /** 한국인 성우 미리보기: 지시문 문구가 있는 단원만 노출 (고르시오, 대화를 듣고, 다음을 듣고, 적절한 것은, 아닌 것/않은 것) */
  function hasInstructionPhrase(lines) {
    if (!lines || !lines.length) return false;
    var text = lines.join(' ');
    return /고르시오/.test(text) || /대화를\s*듣고/.test(text) || /다음을\s*듣고/.test(text) || /적절한\s*것은/.test(text) || /않은\s*것은/.test(text) || /아닌\s*것은/.test(text);
  }

  /** 외국인 성우 미리보기: 대화 지문(W:, M:)이 있는 단원만 노출 */
  function hasDialogueWM(lines) {
    if (!lines || !lines.length) return false;
    var text = lines.join(' ');
    return /\bW\s*:\s*/.test(text) && /\bM\s*:\s*/.test(text);
  }

  /** 목차 페이지에서 '영어듣기 모의고사' + p.N 줄을 파싱해 페이지 번호 -> '## 영어듣기 모의고사' 맵 반환 */
  function parseTocFromPages(pagesData) {
    var map = {};
    if (!pagesData || !pagesData.length) return map;
    for (var i = 0; i < pagesData.length; i++) {
      var p = pagesData[i];
      var text = (p.lines || []).join(' ');
      if (!/목차|차례/.test(text) || !/영어듣기\s*모의고사/i.test(text)) continue;
      var lines = p.lines || [];
      for (var j = 0; j < lines.length; j++) {
        var line = (lines[j] || '').trim();
        var pageMatch = line.match(/p\.\s*(\d+)/i);
        if (!pageMatch) continue;
        var pageNum = parseInt(pageMatch[1], 10);
        var unitNum = null;
        var m1 = line.match(/영어듣기\s*모의고사\s+(\d{1,2})\s+p\./i);
        var m2 = line.match(/^(\d{1,2})\s+영어듣기\s*모의고사\s+p\./i);
        if (m1) unitNum = m1[1];
        else if (m2) unitNum = m2[1];
        if (unitNum != null) {
          var label = (unitNum.length === 1 ? '0' + unitNum : unitNum) + ' 영어듣기 모의고사';
          map[pageNum] = label;
        }
      }
    }
    return map;
  }

  /** 단원명 정규화: '영어듣기 모의고사 ##' -> '## 영어듣기 모의고사' */
  function normalizeUnitLabelForPreview(label) {
    if (!label || !label.trim()) return label;
    var t = label.trim();
    var m = t.match(/영어듣기\s*모의고사\s+(\d{1,2})\s*$/i) || t.match(/^(\d{1,2})\s+영어듣기\s*모의고사\s*$/i);
    if (m) return (m[1].length === 1 ? '0' + m[1] : m[1]) + ' 영어듣기 모의고사';
    return t;
  }

  /** 한국인 성우 미리보기: 단원별 페이지 합치기. 단원 시작 페이지(pageToUnitLabel 있음)부터 다음 단원 시작 전까지 한 묶음, 데이터는 페이지 순으로 이어 붙임. */
  function mergePagesByUnit(pages, pageToUnitLabel) {
    if (!pages || !pages.length) return pages;
    var units = [];
    var currentUnit = null;
    for (var i = 0; i < pages.length; i++) {
      var p = pages[i];
      var label = pageToUnitLabel[p.pageNum];
      if (label) {
        currentUnit = { unitLabel: label, pageStart: p.pageNum, pageEnd: p.pageNum, pages: [p], thumbnailDataUrl: p.thumbnailDataUrl };
        units.push(currentUnit);
      } else if (currentUnit) {
        currentUnit.pageEnd = p.pageNum;
        currentUnit.pages.push(p);
      } else {
        currentUnit = { unitLabel: null, pageStart: p.pageNum, pageEnd: p.pageNum, pages: [p], thumbnailDataUrl: p.thumbnailDataUrl };
        units.push(currentUnit);
      }
    }
    return units.map(function (u) {
      var mergedLines = [];
      for (var j = 0; j < u.pages.length; j++) {
        var lns = u.pages[j].lines || [];
        if (mergedLines.length) mergedLines.push('');
        mergedLines = mergedLines.concat(lns);
      }
      var pageRange = u.pageStart === u.pageEnd ? String(u.pageStart) : (u.pageStart + '~' + u.pageEnd);
      return {
        pageNum: pageRange,
        lines: mergedLines,
        thumbnailDataUrl: u.thumbnailDataUrl,
        unitLabel: u.unitLabel
      };
    });
  }

  /** 퍼플렉시티 제안: 번호 패턴(01~17, \[16-17\]) 기준 구간 나누기 + 옵션 제거 + 끝맺음 정규화. 단원 내 등장 순서 유지(같은 번호 여러 번 허용). */
  function extractInstructionsPerplexityStyle(lines) {
    if (!lines || !lines.length) return [];
    var numOnlyRe = /^\s*(0[1-9]|1[0-7])\s*$/;
    var numStartRe = /^\s*(0[1-9]|1[0-7])\s+(.+)$/;
    var bracketRe = /^\s*\[16-17\]\s*$/;
    var bracketStartRe = /^\s*\[16-17\]\s+(.+)$/;
    var result = [];
    var currentNum = null;
    var accumulated = [];

    function flushInstruction() {
      if (currentNum == null || !accumulated.length) return;
      var text = accumulated.join(' ').replace(/\s{2,}/g, ' ').trim();
      text = removeOptionSegments(text);
      text = normalizeInstructionEnd(text);
      if (text) result.push(currentNum + ' ' + text);
      accumulated = [];
    }
    function removeOptionSegments(s) {
      if (!s) return s;
      var t = s
        .replace(/\$\d+/g, '')
        .replace(/[①②③④]\s*/g, '')
        .replace(/\b[1-4]\)\s*/g, '');
      return t.replace(/\s{2,}/g, ' ').trim();
    }
    function normalizeInstructionEnd(s) {
      if (!s) return s;
      s = s.trim();
      if (/고르시오\.?\s*$/.test(s)) return s.replace(/고르시오\.?\s*$/, '고르시오.');
      if (/고르\s*$/.test(s)) return s.replace(/고르\s*$/, '고르시오.');
      return s;
    }

    for (var i = 0; i < lines.length; i++) {
      var raw = (lines[i] || '').trim();
      if (!raw) continue;
      if (bracketRe.test(raw)) {
        flushInstruction();
        currentNum = '[16-17]';
        accumulated = [];
        continue;
      }
      var bracketMatch = raw.match(bracketStartRe);
      if (bracketMatch) {
        flushInstruction();
        currentNum = '[16-17]';
        accumulated = [bracketMatch[1].trim()];
        continue;
      }
      if (numOnlyRe.test(raw)) {
        flushInstruction();
        currentNum = raw.match(/^\s*(0[1-9]|1[0-7])/)[1];
        accumulated = [];
        continue;
      }
      var numMatch = raw.match(numStartRe);
      if (numMatch) {
        flushInstruction();
        currentNum = numMatch[1];
        accumulated = [numMatch[2].trim()];
        continue;
      }
      if (currentNum != null) {
        accumulated.push(raw);
      }
    }
    flushInstruction();
    return result;
  }

  /** 한국인 성우 미리보기: 퍼플렉시티 방식 지시문 추출 시도 후, 추출 결과가 있으면 사용하고 없으면 띄어쓰기만 통일. */
  function cleanPreviewInstructionLines(lines) {
    if (!lines || !lines.length) return lines || [];
    var extracted = extractInstructionsPerplexityStyle(lines);
    if (extracted && extracted.length > 0) return extracted;
    return lines.map(function (line) { return (line || '').replace(/\s{2,}/g, ' '); });
  }

  /** 한국인 성우 미리보기: 단원명 아래 3줄/4줄 블록이면 첫 줄(단원명)과 마지막 줄(숫자쌍)만 보존 */
  function collapseUnitHeaderBlock(lines) {
    if (!lines || !lines.length) return lines;
    var unitHeaderRe = /^\d{1,2}\s+영어듣기\s*모의고사\s*$/;
    var numberPairRe = /^\d{1,2}\s+\d{1,2}\s*$/;
    var result = [];
    var i = 0;
    while (i < lines.length) {
      var t = (lines[i] || '').trim();
      if (unitHeaderRe.test(t)) {
        if (i + 2 < lines.length && numberPairRe.test((lines[i + 2] || '').trim())) {
          result.push(lines[i]);
          result.push(lines[i + 2]);
          i += 3;
          continue;
        }
        if (i + 3 < lines.length && numberPairRe.test((lines[i + 3] || '').trim())) {
          result.push(lines[i]);
          result.push(lines[i + 3]);
          i += 4;
          continue;
        }
      }
      result.push(lines[i]);
      i++;
    }
    return result;
  }

  /** 미리보기 영역 표시용: 정답 및 해설, 안내 지시문, 페이지 하단(#회·페이지번호, indd·날짜, 수능만만 기본 영어듣기 모의고사) 제외 */
  function filterPreviewLines(lines) {
    if (!lines || !lines.length) return lines;
    var excludePatterns = [
      /정답\s*및\s*해설/,
      /\d+번\s*부터\s*\d+번\s*까지.*듣고\s*답하는/,
      /한\s*번만\s*들려주고|두\s*번\s*들려줍니다/,
      /방송을\s*잘\s*듣고\s*답을\s*하기\s*바랍니다/,
      /녹음을\s*다시\s*한\s*번\s*듣고/,
      /빈칸에\s*알맞은\s*말을\s*쓰시오/,
      /^\s*\d+\s+수능만만\s+(기본\s+)?영어듣기\s*모의고사/,
      /\.indd\s+\d+\s+\d{4}-\d{2}-\d{2}/,
      /^\s*\d{4}-\d{2}-\d{2}\s+(오전|오후)\s+\d/,
      /수능만만.*\.indd.*\d{4}-\d{2}-\d{2}/,
      /\d+\s*회\s+\d+(\s+\d+)*\s*$/,
      /기본\s*듣기모의\s*\d+회\s*\(본문\)\s*\.indd/,
      /\d{4}\.\s*\d{1,2}\.\s*\d{1,2}\..*(오전|오후)\s*\d{1,2}:\d{2}/,
      /(오전|오후)\s*\d{1,2}:\d{2}\s*(오전|오후)\s*\d{1,2}:\d{2}/
    ];
    return lines.filter(function (line) {
      var t = (line || '').trim();
      if (!t) return true;
      for (var i = 0; i < excludePatterns.length; i++) {
        if (excludePatterns[i].test(t)) return false;
      }
      return true;
    });
  }

  function isFirstUnitStart(unitHeader, line, prevLine, nextLine) {
    var name = (unitHeader.name || '').trim();
    var num = (unitHeader.num || '').trim();
    if (/영어듣기|모의고사|듣기\s*모의|DICTATION|Words?\s*[&와]?\s*Expressions?/i.test(name)) return true;
    var combined = (line || '').trim();
    if (prevLine) combined = (prevLine + ' ' + combined).trim();
    if (nextLine) combined = (combined + ' ' + nextLine).trim();
    if (/01\s*영어듣기|01\s*DICTATION|01\s*Word|제\s*1\s*단원|1\s*단원\s*영어/i.test(combined)) return true;
    if ((num === '1' || num === '01') && name.length > 0) return true;
    return false;
  }

  function isMetadataOrFooter(line) {
    if (!line || !line.trim()) return true;
    var t = line.trim();
    if (/\.indd\s|오후\s*\d|오전\s*\d|\d{4}-\d{2}-\d{2}|_본문\s*\(|\.pdf\s*$/i.test(t)) return true;
    if (/^\d+\s*수능만만|회_본문|최종\.indd/.test(t)) return true;
    return false;
  }

  function cleanUnitName(name) {
    if (!name) return name;
    return name.replace(/\s*p\.\s*\d+.*$/gi, '').replace(/\s*p\.\s*\d+/g, '').trim() || name;
  }

  function isValidUnitName(name) {
    var c = cleanUnitName(name);
    return c && !/^p\.\s*\d+$/i.test(c) && !/^\d+\s*$/.test(c);
  }

  function isUnitTitleLine(line) {
    if (!line || !line.trim()) return false;
    var t = line.trim();
    return /영어듣기\s*모의고사|D\s*ICTATION|DICTATION|Words?\s*[&와]?\s*Expressions?/i.test(t) && !isMetadataOrFooter(t);
  }

  /** 상단 큰 글씨 단원 헤더 비교용: 영어듣기 모의고사 / DICTATION / Word & Expressions 통일 */
  function normalizeUnitNameForCompare(name) {
    if (!name || !String(name).trim()) return '';
    var t = String(name).trim().replace(/\s+/g, ' ');
    if (/Words?\s*[&와]?\s*Expressions?/i.test(t)) return 'Word & Expressions';
    if (/DICTATION/i.test(t)) return 'DICTATION';
    if (/영어듣기\s*모의고사/.test(t)) return '영어듣기 모의고사';
    return t;
  }

  function detectUnitHeader(line, prevLine, nextLine) {
    var combined = line.trim();
    if (!combined || isMetadataOrFooter(combined)) return null;
    var m;
    m = combined.match(/정답\s*및\s*해설\s*(\d{1,2})\s*p\./i);
    if (m && prevLine != null && prevLine.trim() && isUnitTitleLine(prevLine)) return { num: m[1].trim(), name: prevLine.trim() };
    if (/\bp\.\s*\d+/.test(combined)) return null;
    if (prevLine != null && isMetadataOrFooter(prevLine)) prevLine = null;
    if (nextLine != null && isMetadataOrFooter(nextLine)) nextLine = null;
    var m = combined.match(/^(\d{1,2})\s*영어듣기\s*모의고사\s*(.*)/);
    if (m) {
      var name = cleanUnitName((m[2] || '영어듣기 모의고사').trim()) || '영어듣기 모의고사';
      return { num: m[1].trim(), name: name };
    }
    m = combined.match(/영어듣기\s*모의고사\s*(\d{1,2})\s*$/);
    if (m) return { num: m[1].trim(), name: '영어듣기 모의고사' };
    m = combined.match(/DICTATION\s*(\d{1,2})\s*$/i);
    if (m) return { num: m[1].trim(), name: 'DICTATION' };
    m = combined.match(/D\s*ICTATION\s*(\d{1,2})\s*$/i);
    if (m) return { num: m[1].trim(), name: 'DICTATION' };
    m = combined.match(/^(\d{1,2})\s*DICTATION\s*(.*)/i);
    if (m) return { num: m[1].trim(), name: cleanUnitName((m[2] || 'DICTATION').trim()) || 'DICTATION' };
    m = combined.match(/^(\d{1,2})\s*Words?\s*[&와]?\s*Expressions?\s*(.*)/i);
    if (m) return { num: m[1].trim(), name: cleanUnitName((m[2] || 'Word & Expressions').trim()) || 'Word & Expressions' };
    m = combined.match(/Words?\s*[&와]?\s*Expressions?\s*(\d{1,2})\s*$/i);
    if (m) return { num: m[1].trim(), name: 'Word & Expressions' };
    m = combined.match(/제?\s*(\d+)\s*단원\s*(.*)/) || combined.match(/Unit\s*(\d+)\s*(.*)/i);
    if (m) {
      var n = cleanUnitName((m[2] || '').trim());
      if (!n || isValidUnitName(n)) return { num: m[1].trim(), name: n || '' };
    }
    m = combined.match(/^(\d+)\s*단원\s*(.*)/);
    if (m) {
      var n2 = cleanUnitName((m[2] || '').trim());
      if (!n2 || isValidUnitName(n2)) return { num: m[1].trim(), name: n2 || '' };
    }
    m = combined.match(/단원\s*(\d+)\s*(.*)/);
    if (m) {
      var n3 = cleanUnitName((m[2] || '').trim());
      if (!n3 || isValidUnitName(n3)) return { num: m[1].trim(), name: n3 || '' };
    }
    var m2 = combined.match(/^(\d+)\s*[\.\s]\s*(.+)/);
    if (m2 && /단원|Unit/.test(combined)) {
      var n4 = cleanUnitName(m2[2].trim());
      if (n4 && isValidUnitName(n4)) return { num: m2[1], name: n4 };
    }
    if (/^\s*\d{1,2}\s*$/.test(combined)) {
      if (nextLine != null && isUnitTitleLine(nextLine)) return { num: combined, name: nextLine.trim() };
      if (prevLine != null && isUnitTitleLine(prevLine)) return { num: combined, name: prevLine.trim() };
    }
    if (isUnitTitleLine(combined) && prevLine != null && /^\s*\d{1,2}\s*$/.test(prevLine.trim()) && !/\bp\.\s*\d+/.test(prevLine + line))
      return { num: prevLine.trim(), name: cleanUnitName(combined) };
    if (prevLine != null) {
      combined = (prevLine + ' ' + line).trim();
      if (/\bp\.\s*\d+/.test(combined)) return null;
      if (!isMetadataOrFooter(combined)) {
        m = combined.match(/^(\d{1,2})\s*영어듣기\s*모의고사\s*(.*)/);
        if (m) return { num: m[1].trim(), name: cleanUnitName((m[2] || '영어듣기 모의고사').trim()) || '영어듣기 모의고사' };
        m = combined.match(/제?\s*(\d+)\s*단원\s*(.*)/) || combined.match(/^(\d+)\s*단원\s*(.*)/);
        if (m) return { num: m[1].trim(), name: cleanUnitName((m[2] || '').trim()) };
        m = combined.match(/^(\d{1,2})\s*DICTATION/i) || combined.match(/^(\d{1,2})\s*Words?\s*[&와]?\s*Expressions?/i);
        if (m) return { num: m[1].trim(), name: cleanUnitName(combined.replace(/^\d{1,2}\s*/, '').trim()) || '영어듣기 모의고사' };
        m = combined.match(/영어듣기\s*모의고사\s*(\d{1,2})\s*$/);
        if (m) return { num: m[1].trim(), name: '영어듣기 모의고사' };
        m = combined.match(/DICTATION\s*(\d{1,2})\s*$/i);
        if (m) return { num: m[1].trim(), name: 'DICTATION' };
        if (/^제?\s*\d+\s*$/.test(prevLine.trim()) && /단원\s*(.*)/.test(line)) {
          var numMatch = prevLine.match(/(\d+)/);
          if (numMatch) return { num: numMatch[1], name: cleanUnitName((line.replace(/단원\s*/, '').trim() || '').trim()) };
        }
      }
    }
    if (nextLine != null && /^\s*\d+\s*$/.test(combined) && !/\bp\.\s*\d+/.test(nextLine) && (/영어듣기\s*모의고사|DICTATION|Words?\s*[&와]?\s*Expressions?/i.test(nextLine) && !isMetadataOrFooter(nextLine))) {
      var numM = line.match(/(\d+)/);
      if (numM) return { num: numM[1], name: cleanUnitName(nextLine.trim()) };
    }
    return null;
  }

  function isQuestionNumber(line) {
    return /^\d+[\.\)]\s*/.test(line.trim()) || /^[①②③④⑤]\s*/.test(line.trim());
  }

  function getQuestionNum(line) {
    const m = line.match(/^(\d+)[\.\)]\s*/) || line.match(/^([①②③④⑤])\s*/);
    return m ? m[1] : null;
  }

  function isWM(line) {
    const t = line.trim();
    return /^W\s*:\s*/.test(t) || /^M\s*:\s*/.test(t);
  }

  function isChoiceLine(line) {
    return /^[①②③④⑤]\s*/.test(line.trim());
  }

  function isQuestionNumberLine(line) {
    return /^\d{2}\s/.test(line.trim()) || /^\d{2}$/.test(line.trim());
  }

  function isInstructionStart(line) {
    var t = line.trim();
    return /다음을\s*듣고|대화를\s*듣고|고르시오|않은\s*것은|적절한\s*것/.test(t);
  }

  /** 지시문 종료: 마침표(.) 또는 물음표(?)로 한 문장이 끝남 */
  function splitInstructionEnd(line) {
    var idx = line.search(/[.?]/);
    if (idx < 0) return { instruction: line, passageStart: null };
    return {
      instruction: line.substring(0, idx + 1).trim(),
      passageStart: line.substring(idx + 1).trim() || null
    };
  }

  function isUnitIntroInstruction(line) {
    var t = line.trim();
    return /^\d+번\s*부터\s*\d+번\s*까지|방송을\s*잘\s*들어|한\s*번만\s*들려주고|두\s*번\s*들려줍니다/.test(t);
  }

  function parsePagesForKoreanDetailed(pagesLines) {
    var result = [];
    var currentUnit = null;
    var currentUnitName = '';
    var currentQuestions = [];
    var state = 'intro';
    var skippingUnitIntro = false;
    var instructionBuf = [];
    var passageBuf = [];
    var choicesBuf = [];
    var lastQNum = null;
    var digitDot = /^(\d+)[\.\)]\s*/;
    var digitTwo = /^(\d{2})\s/;
    var circleNum = /^([①②③④⑤])\s*/;

    for (var p = 0; p < pagesLines.length; p++) {
      var lines = pagesLines[p];
      if (!lines.length) continue;
      if (isCoverOrToc(lines, p)) continue;

      for (var i = 0; i < lines.length; i++) {
        var line = lines[i];
        var prevLine = i > 0 ? lines[i - 1] : null;
        var nextLine = i < lines.length - 1 ? lines[i + 1] : null;
        var unitHeader = detectUnitHeader(line, prevLine, nextLine);
        if (unitHeader) {
          if (lastQNum !== null && (instructionBuf.length || passageBuf.length || choicesBuf.length)) {
            currentQuestions.push({
              questionNo: lastQNum,
              instruction: instructionBuf.join(' ').trim(),
              passage: passageBuf.join(' ').trim(),
              choices: choicesBuf.join('\n').trim()
            });
          }
          if (currentUnit !== null && currentQuestions.length) {
            result.push({ unitNo: currentUnit, unitName: currentUnitName, questions: currentQuestions.slice() });
          }
          currentUnit = unitHeader.num;
          currentUnitName = unitHeader.name || currentUnitName;
          currentQuestions = [];
          instructionBuf = [];
          passageBuf = [];
          choicesBuf = [];
          lastQNum = null;
          state = 'intro';
          skippingUnitIntro = true;
          continue;
        }
        if (currentUnit === null) continue;
        if (skippingUnitIntro) {
          if (isUnitIntroInstruction(line)) continue;
          if (isQuestionNumberLine(line) || isInstructionStart(line)) skippingUnitIntro = false;
          else continue;
        }

        var qMatch = line.match(digitDot) || line.match(digitTwo);
        if (qMatch) {
          if (lastQNum !== null) {
            currentQuestions.push({
              questionNo: lastQNum,
              instruction: instructionBuf.join(' ').trim(),
              passage: passageBuf.join(' ').trim(),
              choices: choicesBuf.join('\n').trim()
            });
          }
          lastQNum = String(parseInt(qMatch[1], 10));
          instructionBuf = [];
          passageBuf = [];
          choicesBuf = [];
          var rest = line.replace(digitDot, '').replace(digitTwo, '').trim();
          if (rest && !isWM(rest) && !isChoiceLine(rest)) instructionBuf.push(rest);
          state = 'instruction';
          continue;
        }
        if (isWM(line)) {
          if (state === 'instruction' && instructionBuf.length) state = 'passage';
          passageBuf.push(line);
          continue;
        }
        if (isChoiceLine(line)) {
          state = 'choices';
          choicesBuf.push(line);
          continue;
        }
        if (state === 'instruction') instructionBuf.push(line);
        else if (state === 'passage' && passageBuf.length) passageBuf.push(line);
        else if (state === 'choices') choicesBuf.push(line);
        else if (state === 'intro' && line.trim() && !currentUnitName) currentUnitName = line.trim();
      }
    }
    if (lastQNum !== null) {
      currentQuestions.push({
        questionNo: lastQNum,
        instruction: instructionBuf.join(' ').trim(),
        passage: passageBuf.join(' ').trim(),
        choices: choicesBuf.join('\n').trim()
      });
    }
    if (currentUnit !== null && currentQuestions.length) {
      result.push({ unitNo: currentUnit, unitName: currentUnitName, questions: currentQuestions });
    }
    return result;
  }

  function parsePagesForAll(pagesLines) {
    const result = [];
    let currentUnit = null;
    let currentUnitName = '';
    let currentLines = [];
    let extractionStarted = false;
    let skippingUnitIntro = false;
    let lastLineOfPrevPage = null;

    for (let p = 0; p < pagesLines.length; p++) {
      const lines = pagesLines[p];
      if (!lines.length) {
        lastLineOfPrevPage = null;
        continue;
      }
      if (isCoverOrToc(lines, p)) {
        lastLineOfPrevPage = lines[lines.length - 1];
        continue;
      }
      if (isExcludedPage(lines, p)) {
        lastLineOfPrevPage = lines[lines.length - 1];
        continue;
      }

      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const prevLine = i > 0 ? lines[i - 1] : lastLineOfPrevPage;
        const nextLine = i < lines.length - 1 ? lines[i + 1] : null;
        if (i === 0) lastLineOfPrevPage = null;
        if (i === lines.length - 1) lastLineOfPrevPage = line;
        const unitHeader = detectUnitHeader(line, prevLine, nextLine);
        if (unitHeader) {
          if (!extractionStarted && !isFirstUnitStart(unitHeader, line, prevLine, nextLine)) continue;
          extractionStarted = true;
          var sameNum = currentUnit !== null && padTwo(unitHeader.num) === padTwo(currentUnit);
          var sameName = normalizeUnitNameForCompare(currentUnitName) === normalizeUnitNameForCompare(unitHeader.name || '');
          if (sameNum && sameName) {
            skippingUnitIntro = true;
            continue;
          }
          if (currentUnit !== null && currentLines.length) {
            result.push({ unitNo: currentUnit, unitName: currentUnitName, lines: currentLines.slice() });
          }
          currentUnit = unitHeader.num;
          currentUnitName = unitHeader.name || currentUnitName;
          currentLines = [line];
          skippingUnitIntro = true;
          continue;
        }
        if (extractionStarted && currentUnit !== null && /^\s*\d{1,2}\s*$/.test(line.trim())) {
          const num = parseInt(line.trim(), 10);
          const curNum = parseInt(currentUnit, 10);
          if (!isNaN(num) && !isNaN(curNum) && num === curNum + 1) {
            const nextIsUnitTitle = nextLine != null && isUnitTitleLine(nextLine);
            result.push({ unitNo: currentUnit, unitName: currentUnitName, lines: currentLines.slice() });
            currentUnit = padTwo(line.trim());
            currentUnitName = nextIsUnitTitle ? nextLine.trim() : '영어듣기 모의고사';
            currentLines = nextIsUnitTitle ? [line, nextLine] : [line];
            skippingUnitIntro = true;
            if (nextIsUnitTitle) i++;
            continue;
          }
        }
        if (currentUnit !== null) {
          if (skippingUnitIntro) {
            if (isQuestionNumberLine(line) || isInstructionStart(line)) {
              skippingUnitIntro = false;
              currentLines.push(line);
            }
            continue;
          }
          currentLines.push(line);
        }
      }
    }
    if (currentUnit !== null && currentLines.length) {
      result.push({ unitNo: currentUnit, unitName: currentUnitName, lines: currentLines });
    }
    if (result.length === 0 && pagesLines.length) {
      var allLines = [];
      for (var pi = 0; pi < pagesLines.length; pi++) {
        if (isCoverOrToc(pagesLines[pi], pi) || isExcludedPage(pagesLines[pi], pi)) continue;
        for (var li = 0; li < pagesLines[pi].length; li++) {
          if (!isMetadataOrFooter(pagesLines[pi][li])) allLines.push(pagesLines[pi][li]);
        }
      }
      if (allLines.length) result.push({ unitNo: '1', unitName: '전체', lines: allLines });
    }
    return result;
  }

  function parsePagesForKorean(allPageLines) {
    const result = [];
    let currentUnit = null;
    let currentUnitName = '';
    let inUnit = false;
    let skippingUnitIntro = false;
    let lastQuestionNum = null;
    let instructionBuffer = [];

    for (let p = 0; p < allPageLines.length; p++) {
      const lines = allPageLines[p];
      if (!lines.length) continue;
      if (isCoverOrToc(lines, p)) continue;

      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const prevLine = i > 0 ? lines[i - 1] : null;
        const nextLine = i < lines.length - 1 ? lines[i + 1] : null;
        const unitHeader = detectUnitHeader(line, prevLine, nextLine);
        if (unitHeader && (unitHeader.name || currentUnit === null)) {
          if (currentUnit !== null && instructionBuffer.length) {
            const content = instructionBuffer.join(' ').trim();
            if (lastQuestionNum && content)
              result.push({ unitNo: currentUnit, unitName: currentUnitName, questionNo: lastQuestionNum, content: content });
          }
          currentUnit = unitHeader.num;
          currentUnitName = unitHeader.name || currentUnitName;
          inUnit = true;
          skippingUnitIntro = true;
          instructionBuffer = [];
          lastQuestionNum = null;
          continue;
        }

        if (!inUnit) continue;
        if (skippingUnitIntro) {
          if (isUnitIntroInstruction(line)) continue;
          if (isQuestionNumberLine(line) || isInstructionStart(line)) skippingUnitIntro = false;
          else continue;
        }

        var twoMatch = line.match(/^(\d{2})\s/);
        var qNum = getQuestionNum(line) || (twoMatch ? twoMatch[1] : null);
        if (qNum !== null) {
          if (lastQuestionNum !== null && instructionBuffer.length) {
            const content = instructionBuffer.join(' ').trim();
            if (content) result.push({ unitNo: currentUnit, unitName: currentUnitName, questionNo: lastQuestionNum, content: content });
          }
          lastQuestionNum = qNum;
          instructionBuffer = [];
          var rest = line.replace(/^\d+[\.\)]\s*/, '').replace(/^\d{2}\s*/, '').replace(/^[①②③④⑤]\s*/, '').trim();
          if (rest && !isWM(rest)) instructionBuffer.push(rest);
          continue;
        }

        if (lastQuestionNum !== null) {
          if (isWM(line)) {
            var content = instructionBuffer.join(' ').trim();
            if (content) result.push({ unitNo: currentUnit, unitName: currentUnitName, questionNo: lastQuestionNum, content: content });
            instructionBuffer = [];
            lastQuestionNum = null;
          } else {
            instructionBuffer.push(line);
          }
        } else if (instructionBuffer.length === 0 && line.trim() && !isWM(line)) {
          if (currentUnitName === '' && currentUnit) currentUnitName = line.trim();
          else instructionBuffer.push(line);
        }
      }
    }

    if (currentUnit !== null && lastQuestionNum && instructionBuffer.length) {
      const content = instructionBuffer.join(' ').trim();
      if (content) result.push({ unitNo: currentUnit, unitName: currentUnitName, questionNo: lastQuestionNum, content: content });
    }
    return result;
  }

  function parsePagesForForeign(allPageLines) {
    const result = [];
    let currentUnit = null;
    let currentUnitName = '';
    let inUnit = false;
    let lastQuestionNum = null;
    let wmBuffer = [];
    let lastQFromWM = null;

    for (let p = 0; p < allPageLines.length; p++) {
      const lines = allPageLines[p];
      if (!lines.length) continue;
      if (isCoverOrToc(lines, p)) continue;

      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const prevLine = i > 0 ? lines[i - 1] : null;
        const nextLine = i < lines.length - 1 ? lines[i + 1] : null;
        const unitHeader = detectUnitHeader(line, prevLine, nextLine);
        if (unitHeader && (unitHeader.name || currentUnit === null)) {
          if (lastQFromWM !== null && wmBuffer.length) {
            result.push({ unitNo: currentUnit, unitName: currentUnitName, questionNo: lastQFromWM, content: wmBuffer.join(' ').trim() });
          }
          currentUnit = unitHeader.num;
          currentUnitName = unitHeader.name || currentUnitName;
          inUnit = true;
          wmBuffer = [];
          lastQuestionNum = null;
          lastQFromWM = null;
          continue;
        }

        if (!inUnit) continue;

        const qNum = getQuestionNum(line);
        if (qNum !== null) {
          if (lastQFromWM !== null && wmBuffer.length)
            result.push({ unitNo: currentUnit, unitName: currentUnitName, questionNo: lastQFromWM, content: wmBuffer.join(' ').trim() });
          lastQuestionNum = qNum;
          lastQFromWM = qNum;
          wmBuffer = [];
          const rest = line.replace(/^\d+[\.\)]\s*/, '').replace(/^[①②③④⑤]\s*/, '').trim();
          if (isWM(rest)) wmBuffer.push(rest);
          continue;
        }

        if (isWM(line)) {
          if (lastQuestionNum) lastQFromWM = lastQuestionNum;
          wmBuffer.push(line);
        } else if (wmBuffer.length) {
          if (lastQFromWM !== null) {
            result.push({ unitNo: currentUnit, unitName: currentUnitName, questionNo: lastQFromWM, content: wmBuffer.join(' ').trim() });
            wmBuffer = [];
            lastQFromWM = null;
          }
        } else if (currentUnitName === '' && currentUnit && line.trim()) {
          currentUnitName = line.trim();
        }
      }
    }

    if (lastQFromWM !== null && wmBuffer.length)
      result.push({ unitNo: currentUnit, unitName: currentUnitName, questionNo: lastQFromWM, content: wmBuffer.join(' ').trim() });
    return result;
  }

  function setLoading(show, percent, text) {
    if (!show) {
      loadingWrap.classList.remove('show');
      return;
    }
    loadingWrap.classList.add('show');
    loadingBarFill.style.width = (percent || 0) + '%';
    loadingText.textContent = text || '';
  }

  async function runExtraction() {
    const file = pdfInput.files[0];
    if (!file) return;
    btnExtract.disabled = true;
    setStatus('', 'info');
    setLoading(true, 0, 'PDF 로딩 중...');

    try {
      const arrayBuffer = await file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;
      const numPages = pdf.numPages;
      const maxPages = (MAX_EXTRACT_PAGES > 0) ? Math.min(numPages, MAX_EXTRACT_PAGES) : numPages;
      const pagesLines = [];
      pagesData = [];

      for (let i = 1; i <= maxPages; i++) {
        const page = await pdf.getPage(i);
        const items = await getPageTextItems(page);
        const viewport = page.getViewport({ scale: 1 });
        const centerX = viewport.width / 2;
        // 항상 2단 구성: 왼쪽 영역(위→아래) 먼저 추출, 이어서 오른쪽 영역(위→아래) 추출
        const lineResults = itemsToLines(items, centerX);
        const lines = addHeaderFooterLabels(lineResults, viewport.height);
        pagesLines.push(lines);
        const thumbnailDataUrl = await renderPageToThumbnailDataUrl(page, 120);
        pagesData.push({ pageNum: i, lines: lines, thumbnailDataUrl: thumbnailDataUrl });
        const pct = Math.round((i / maxPages) * 100);
        setLoading(true, pct, '페이지 ' + i + ' / ' + maxPages + ' 분석 중...');
      }

      allPageLines = pagesLines;
      pageToUnitLabel = parseTocFromPages(pagesData);
      // 추출 방식(목차 1단·그 외 2단, [머리글]/[바닥글] 구분)은 한국인·외국인 성우 동일하게 적용됨(pagesData 공유).
      // 한국인 성우: 미리보기에 표지·구성·목차 제외, 지시문 없는 페이지만. 외국인 성우: 전체 페이지.
      previewPagesKorean = pagesData.filter(function (p) {
        return !isExcludedFromPreview(p.lines) && hasInstructionPhrase(p.lines);
      });
      previewPagesForeign = pagesData.slice();
      renderPagePreviewPanel(pagesData, scriptType === 'korean' ? previewPagesKorean : previewPagesForeign, scriptType);
      setLoading(true, 100, '데이터 정리 중...');
      if (typeof console !== 'undefined' && console.log) {
        var sampleFrom = Math.min(3, Math.max(0, maxPages - 10));
        var sampleTo = Math.min(sampleFrom + 8, maxPages);
        console.log('=== PDF 줄 패턴 (페이지 ' + (sampleFrom + 1) + '~' + sampleTo + ') ===');
        for (var pi = sampleFrom; pi < sampleTo; pi++) {
          var pl = pagesLines[pi] || [];
          console.log('페이지 ' + (pi + 1) + ' (' + pl.length + '줄):', pl.slice(0, 15).map(function (l) { return (l || '').substring(0, 60); }));
        }
        console.log('=== 끝 ===');
      }
      allUnitsData = parsePagesForAll(allPageLines);
      allUnitsData.sort(function (a, b) {
        var na = parseInt(a.unitNo, 10);
        var nb = parseInt(b.unitNo, 10);
        if (!isNaN(na) && !isNaN(nb)) return na - nb;
        return String(a.unitNo).localeCompare(String(b.unitNo));
      });
      var merged = [];
      for (var u = 0; u < allUnitsData.length; u++) {
        var unit = allUnitsData[u];
        var last = merged.length ? merged[merged.length - 1] : null;
        var sameAsLast = last && padTwo(unit.unitNo) === padTwo(last.unitNo) && normalizeUnitNameForCompare(unit.unitName) === normalizeUnitNameForCompare(last.unitName);
        if (sameAsLast && last.lines) {
          last.lines = last.lines.concat(unit.lines || []);
        } else {
          merged.push({ unitNo: unit.unitNo, unitName: unit.unitName, lines: (unit.lines || []).slice() });
        }
      }
      allUnitsData = merged;
      allUnitsData.forEach(function (unit) {
        unit.questions = parseUnitLinesToQuestions(unit.lines || []);
      });
      allUnitsDataKorean = parsePagesForKoreanDetailed(pagesLines);
      extractedData = scriptType === 'korean'
        ? parsePagesForKorean(allPageLines)
        : parsePagesForForeign(allPageLines);

      setLoading(false);
      setStatus('추출 완료', 'success');
      setFontAndPreviewAreasEnabled(true);
      renderUnitsOnly();
      renderAccordion();
    } catch (err) {
      setLoading(false);
      setStatus('오류: ' + (err.message || String(err)), 'error');
      console.error(err);
    }
    btnExtract.disabled = false;
  }

  function padTwo(n) {
    var s = String(n);
    return s.length >= 2 ? s : '0' + s;
  }

  function isWordListLine(line) {
    var t = line.trim();
    if (!t) return false;
    if (/□/.test(t)) return true;
    if (/^[a-zA-Z\s\-']+[\s\u00A0]+[가-힣\s]+$/.test(t) || /[가-힣][\s\u00A0]+[a-zA-Z]/.test(t)) return true;
    return false;
  }

  var MAX_QUESTIONS_PER_UNIT = 17;

  function parseUnitLinesToQuestions(lines) {
    var questions = [];
    var digitDot = /^(\d+)[\.\)]\s*/;
    var digitTwoWithContent = /^(\d{2})\s+/;
    var state = 'skip';
    var instructionBuf = [];
    var passageBuf = [];
    var choicesBuf = [];
    var wordsBuf = [];
    var lastQNum = null;
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];
      if (isUnitIntroInstruction(line)) continue;
      if (i === 0 && /영어듣기\s*모의고사|DICTATION|Words?\s*[&와]?\s*Expressions?/i.test(line)) continue;
      if (i === 0 && /^\d{1,2}\s*$/.test(line.trim())) continue;
      var lineTrim = line.trim();
      var twoNumsOnly = /^\d{1,2}\s+\d{1,2}\s*$/.test(lineTrim);
      if (twoNumsOnly) {
        if (state === 'instruction') passageBuf.push(line);
        else if (state === 'passage') passageBuf.push(line);
        else if (state === 'choices') choicesBuf.push(line);
        continue;
      }
      var qMatch = line.match(digitDot) || line.match(digitTwoWithContent);
      var digitOnly = lineTrim.match(/^(\d{1,2})$/);
      if (digitOnly && !qMatch) {
        var numOnly = parseInt(digitOnly[1], 10);
        var nextExpected = lastQNum === null ? 1 : parseInt(lastQNum, 10) + 1;
        if (numOnly >= 1 && numOnly <= MAX_QUESTIONS_PER_UNIT && numOnly === nextExpected) {
          qMatch = digitOnly;
        } else {
          if (state === 'passage') passageBuf.push(line);
          else if (state === 'instruction') passageBuf.push(line);
          else if (state === 'choices') choicesBuf.push(line);
          continue;
        }
      }
      if (qMatch) {
        var num = parseInt(qMatch[1], 10);
        if (num < 1 || num > MAX_QUESTIONS_PER_UNIT) {
          if (state === 'instruction') instructionBuf.push(line);
          else if (state === 'passage') passageBuf.push(line);
          else if (state === 'choices') choicesBuf.push(line);
          continue;
        }
        var numStr = num <= 9 ? '0' + num : String(num);
        var alreadyHave = questions.some(function (q) { return String(q.questionNo) === numStr; });
        if (alreadyHave && digitOnly) {
          if (state === 'passage') passageBuf.push(line);
          else if (state === 'choices') choicesBuf.push(line);
          else if (state === 'instruction') instructionBuf.push(line);
          continue;
        }
        if (questions.length >= MAX_QUESTIONS_PER_UNIT) {
          if (state === 'choices') choicesBuf.push(line);
          continue;
        }
        if (lastQNum !== null) {
          questions.push({
            questionNo: lastQNum,
            instruction: instructionBuf.join(' ').trim(),
            passage: passageBuf.join('\n').trim(),
            choices: choicesBuf.join('\n').trim(),
            words: wordsBuf.join('\n').trim()
          });
        }
        lastQNum = num <= 9 ? '0' + num : String(num);
        instructionBuf = [];
        passageBuf = [];
        choicesBuf = [];
        wordsBuf = [];
        var rest = line.replace(digitDot, '').replace(digitTwoWithContent, '').replace(/^(\d{1,2})\s*$/, '').trim();
        if (rest && !isChoiceLine(rest)) {
          if (isWordListLine(rest)) {
            wordsBuf.push(rest);
          } else {
            var split = splitInstructionEnd(rest);
            if (split.instruction) instructionBuf.push(split.instruction);
            if (split.passageStart) {
              passageBuf.push(split.passageStart);
              state = 'passage';
            } else {
              state = 'instruction';
            }
          }
        } else {
          state = 'instruction';
        }
        continue;
      }
      if (isChoiceLine(line)) {
        state = 'choices';
        choicesBuf.push(line);
        continue;
      }
      if (isWM(line)) {
        if (state === 'instruction') state = 'passage';
        passageBuf.push(line);
        continue;
      }
      if (state === 'instruction') {
        var split = splitInstructionEnd(line);
        if (split.instruction) instructionBuf.push(split.instruction);
        if (split.passageStart) {
          passageBuf.push(split.passageStart);
          state = 'passage';
        }
        continue;
      }
      if (state === 'passage') {
        passageBuf.push(line);
        continue;
      }
      if (state === 'choices') {
        choicesBuf.push(line);
        continue;
      }
      if (isWordListLine(line)) {
        wordsBuf.push(line);
      }
    }
    if (lastQNum !== null && questions.length < MAX_QUESTIONS_PER_UNIT) {
      questions.push({
        questionNo: lastQNum,
        instruction: instructionBuf.join(' ').trim(),
        passage: passageBuf.join('\n').trim(),
        choices: choicesBuf.join('\n').trim(),
        words: wordsBuf.join('\n').trim()
      });
    }
    questions.sort(function (a, b) {
      var na = parseInt(a.questionNo, 10);
      var nb = parseInt(b.questionNo, 10);
      return (na - nb);
    });
    return questions;
  }

  function renderUnitsOnly() {
    if (!unitsOnlyList) return;
    unitsOnlyList.innerHTML = '';
    allUnitsData.forEach(function (unit) {
      var qCount = (unit.questions && unit.questions.length) ? unit.questions.length : parseUnitLinesToQuestions(unit.lines || []).length;
      var nameText = (unit.unitName || '') + (qCount ? ' (' + qCount + '문항)' : '');
      var chip = document.createElement('span');
      chip.className = 'unit-chip';
      chip.innerHTML = '<span class="num">' + escapeHtml(padTwo(unit.unitNo)) + '</span> <span class="name">' + escapeHtml(nameText) + '</span>';
      chip.title = '단원 번호 ' + escapeHtml(unit.unitNo) + ' · 단원명 ' + escapeHtml(unit.unitName || '') + (qCount ? ' · ' + qCount + '문항' : '');
      unitsOnlyList.appendChild(chip);
    });
  }

  function renderAccordion() {
    if (!previewAccordion) return;
    var fontSize = getFontSize();
    previewAccordion.innerHTML = '';
    previewAccordion.style.fontSize = fontSize + 'px';
    allUnitsData.forEach(function (unit, idx) {
      var item = document.createElement('div');
      item.className = 'accordion-item' + (idx === 0 ? ' open' : '');
      var header = document.createElement('div');
      header.className = 'accordion-header';
      var questions = (unit.questions && unit.questions.length) ? unit.questions : parseUnitLinesToQuestions(unit.lines || []);
      var qCount = questions.length;
      var unitTitle = padTwo(unit.unitNo) + (unit.unitName ? ' ' + unit.unitName : '') + (qCount ? ' (' + qCount + '문항)' : '');
      header.innerHTML = '<span>' + escapeHtml(unitTitle) + '</span><span class="arrow">▼</span>';
      var body = document.createElement('div');
      body.className = 'accordion-body';
      var content = document.createElement('div');
      content.className = 'accordion-content';
      if (questions.length) {
        var html = [];
        questions.forEach(function (q) {
          html.push('<div class="preview-question-block">');
          html.push('<div class="preview-q-num">' + escapeHtml(padTwo(q.questionNo)) + '</div>');
          html.push('<div class="preview-block-label">[지시문]</div>');
          html.push('<div class="preview-block-content">' + escapeHtml(q.instruction || '') + '</div>');
          html.push('<div class="preview-block-label">[지문]</div>');
          html.push('<div class="preview-block-content">' + escapeHtml(q.passage || '') + '</div>');
          html.push('<div class="preview-block-label">[선택지]</div>');
          html.push('<div class="preview-block-content">' + escapeHtml(q.choices || '') + '</div>');
          if (q.words) {
            html.push('<div class="preview-block-label">[단어]</div>');
            html.push('<div class="preview-block-content">' + escapeHtml(q.words) + '</div>');
          }
          html.push('</div>');
        });
        content.innerHTML = html.join('');
      } else {
        content.textContent = (unit.lines || []).join('\n');
      }
      body.appendChild(content);
      item.appendChild(header);
      item.appendChild(body);
      header.addEventListener('click', function () {
        item.classList.toggle('open');
      });
      previewAccordion.appendChild(item);
    });
  }

  function escapeHtml(s) {
    if (s == null) return '';
    const div = document.createElement('div');
    div.textContent = s;
    return div.innerHTML;
  }

  function getFirstUnitPreview() {
    if (!extractedData.length) return '';
    const firstUnit = extractedData[0].unitNo;
    const rows = extractedData.filter(function (r) { return r.unitNo === firstUnit; });
    const fontSize = Math.max(10, Math.min(30, parseInt(fontSizeInput.value, 10) || 15));
    const lines = rows.map(function (r) {
      return '[' + r.unitNo + '단원] ' + (r.unitName ? r.unitName + ' | ' : '') + '문항 ' + r.questionNo + '\n' + r.content;
    });
    return { html: lines.join('\n\n'), fontSize: fontSize };
  }

  /** 폰트 미리보기 팝업: 파일명(+5px 볼드), 한줄띄기, 단원1만 단원명(+2px 볼드). korean=문제번호+지시문, foreign=M/W 색상·문항번호 볼드 */
  function buildFontPreviewHtml() {
    var fontSize = getFontSize();
    var f5 = fontSize + 5;
    var f2 = fontSize + 2;
    if (previewExportData && previewExportData.blocks && previewExportData.blocks.length > 0) {
      var fileName = previewExportData.fileName || '';
      var first = previewExportData.blocks[0];
      var unitLabel = first.unitLabel || '';
      var contentLines = first.contentLines || [];
      var parts = [];
      parts.push('<p style="font-size:' + f5 + 'px;font-weight:bold;margin:0 0 .5em 0">' + escapeHtml(fileName) + '</p>');
      parts.push('<p style="margin:0 0 .5em 0"><br></p>');
      if (unitLabel) parts.push('<p style="font-size:' + f2 + 'px;font-weight:bold;margin:0 0 .5em 0">' + escapeHtml(unitLabel) + '</p>');
      if (scriptType === 'foreign') {
        var COLOR_M = '#1565c0';
        var COLOR_W = '#c2185b';
        var COLOR_OTHER = '#2e7d32';
        contentLines.forEach(function (ln) {
          if (ln === '') {
            parts.push('<p style="margin:0 0 .35em 0"><br></p>');
          } else if (isForeignQuestionNumberLine(ln)) {
            parts.push('<p style="font-size:' + fontSize + 'px;font-weight:bold;margin:0 0 .35em 0">' + escapeHtml(ln) + '</p>');
          } else if (/^M:\s/.test(ln)) {
            parts.push('<p style="font-size:' + fontSize + 'px;margin:0 0 .35em 0;color:' + COLOR_M + '">' + escapeHtml(ln) + '</p>');
          } else if (/^W:\s/.test(ln)) {
            parts.push('<p style="font-size:' + fontSize + 'px;margin:0 0 .35em 0;color:' + COLOR_W + '">' + escapeHtml(ln) + '</p>');
          } else {
            parts.push('<p style="font-size:' + fontSize + 'px;margin:0 0 .35em 0;color:' + COLOR_OTHER + '">' + escapeHtml(ln) + '</p>');
          }
        });
      } else {
        contentLines.forEach(function (ln) {
          parts.push('<p style="font-size:' + fontSize + 'px;margin:0 0 .35em 0">' + escapeHtml(ln) + '</p>');
        });
      }
      return parts.join('');
    }
    var first = getFirstUnitPreview();
    if (!first || !first.html) return '<p style="font-size:' + fontSize + 'px">미리보기 데이터가 없습니다. PDF를 분석한 뒤 미리보기 영역에 내용이 표시된 상태에서 확인하세요.</p>';
    return '<p style="font-size:' + fontSize + 'px;white-space:pre-wrap">' + escapeHtml(first.html) + '</p>';
  }

  function showFontPreview() {
    previewContent.style.fontSize = '';
    previewContent.style.fontFamily = 'NanumSquareNeo, sans-serif';
    previewContent.innerHTML = buildFontPreviewHtml();
    previewPopup.classList.add('show');
  }

  function getFilename(ext) {
    const now = new Date();
    const date = now.getFullYear() + '' + String(now.getMonth() + 1).padStart(2, '0') + String(now.getDate()).padStart(2, '0');
    const time = String(now.getHours()).padStart(2, '0') + String(now.getMinutes()).padStart(2, '0') + String(now.getSeconds()).padStart(2, '0');
    const base = (uploadedFileName || 'script') + '_' + date + '_' + time;
    return base + ext;
  }

  function getFontSize() {
    return Math.max(10, Math.min(30, parseInt(fontSizeInput.value, 10) || 15));
  }

  fontSizeInput.addEventListener('change', function () {
    if (allUnitsDataKorean.length || allUnitsData.length) renderAccordion();
  });
  fontSizeInput.addEventListener('input', function () {
    if (allUnitsDataKorean.length || allUnitsData.length) renderAccordion();
  });

  function getFullQuestionRows() {
    if (!allUnitsData || !allUnitsData.length) return [];
    var rows = [];
    allUnitsData.forEach(function (unit) {
      var questions = (unit.questions && unit.questions.length) ? unit.questions : parseUnitLinesToQuestions(unit.lines || []);
      questions.forEach(function (q) {
        rows.push({
          unitNo: unit.unitNo,
          unitName: unit.unitName || '',
          questionNo: q.questionNo || '',
          instruction: q.instruction || '',
          passage: q.passage || '',
          choices: q.choices || '',
          words: q.words || ''
        });
      });
    });
    return rows;
  }

  /** 단원 라벨에서 단원번호·단원명 분리 (예: "01 영어듣기 모의고사" -> { unitNo: "01", unitName: "영어듣기 모의고사" }) */
  function parseUnitLabel(unitLabel) {
    if (!unitLabel || !unitLabel.trim()) return { unitNo: '', unitName: '' };
    var m = (unitLabel.trim()).match(/^(\d{1,2})\s+(.+)$/);
    if (m) return { unitNo: m[1], unitName: m[2] };
    var m2 = (unitLabel.trim()).match(/^기출\s+(\d{1,2})\s+(.+)$/);
    if (m2) return { unitNo: m2[1], unitName: '기출 ' + m2[2] };
    return { unitNo: '', unitName: unitLabel.trim() };
  }

  /** 문제줄 "01 다음을 듣고..." 에서 문항번호·지시문 분리 */
  function parseQuestionLine(line) {
    if (!line || !line.trim()) return { questionNo: '', instruction: '' };
    var m = (line.trim()).match(/^(\d{1,2})\s+(.+)$/);
    return m ? { questionNo: m[1], instruction: m[2] } : { questionNo: '', instruction: line.trim() };
  }

  function downloadFile(type) {
    const fontSize = getFontSize();
    const fontFamily = 'NanumSquareNeo, sans-serif';
    var usePreview = previewExportData && previewExportData.blocks && previewExportData.blocks.length > 0;
    var fullRows = getFullQuestionRows();
    var hasFull = fullRows.length > 0;
    if (!usePreview && !hasFull && !extractedData.length) return;

    if (usePreview) {
      var fileName = previewExportData.fileName || '';
      var blocks = previewExportData.blocks;

      if (type === 'text') {
        var textParts = [fileName, ''];
        blocks.forEach(function (b) {
          if (b.unitLabel) textParts.push(b.unitLabel);
          (b.contentLines || []).forEach(function (ln) { textParts.push(ln); });
          textParts.push('');
        });
        var textOut = textParts.join('\n').replace(/\n{3,}/g, '\n\n').trim();
        if (textOut) textOut += '\n';
        saveBlob(new Blob([textOut], { type: 'text/plain;charset=utf-8' }), getFilename('.txt'));
        return;
      }

      if (type === 'excel') {
        var wsData = [[fileName]];
        blocks.forEach(function (b) {
          var u = parseUnitLabel(b.unitLabel || '');
          wsData.push([u.unitNo, u.unitName, '', '']);
          (b.contentLines || []).forEach(function (ln) {
            var q = parseQuestionLine(ln);
            wsData.push([u.unitNo, u.unitName, q.questionNo, q.instruction]);
          });
        });
        var wb = XLSX.utils.book_new();
        var ws = XLSX.utils.aoa_to_sheet(wsData);
        XLSX.utils.book_append_sheet(wb, ws, '스크립트');
        XLSX.writeFile(wb, getFilename('.xlsx'));
        return;
      }

      if (type === 'word') {
        var f5 = fontSize + 5;
        var f2 = fontSize + 2;
        var bodyParts = ['<p style="font-size:' + f5 + 'px;font-weight:bold">' + escapeHtml(fileName) + '</p>', '<p><br></p>'];
        var COLOR_M = '#1565c0';
        var COLOR_W = '#c2185b';
        var COLOR_OTHER = '#2e7d32';
        blocks.forEach(function (b) {
          if (b.unitLabel) bodyParts.push('<p style="font-size:' + f2 + 'px;font-weight:bold">' + escapeHtml(b.unitLabel) + '</p>');
          (b.contentLines || []).forEach(function (ln) {
            if (scriptType === 'foreign') {
              if (ln === '') bodyParts.push('<p><br></p>');
              else if (isForeignQuestionNumberLine(ln)) bodyParts.push('<p style="font-size:' + fontSize + 'px;font-weight:bold">' + escapeHtml(ln) + '</p>');
              else if (/^M:\s/.test(ln)) bodyParts.push('<p style="font-size:' + fontSize + 'px;color:' + COLOR_M + '">' + escapeHtml(ln) + '</p>');
              else if (/^W:\s/.test(ln)) bodyParts.push('<p style="font-size:' + fontSize + 'px;color:' + COLOR_W + '">' + escapeHtml(ln) + '</p>');
              else bodyParts.push('<p style="font-size:' + fontSize + 'px;color:' + COLOR_OTHER + '">' + escapeHtml(ln) + '</p>');
            } else {
              bodyParts.push('<p style="font-size:' + fontSize + 'px">' + escapeHtml(ln) + '</p>');
            }
          });
          bodyParts.push('<p><br></p>');
        });
        var html = '<!DOCTYPE html><html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word"><head><meta charset="UTF-8"><link href="https://hangeul.pstatic.net/hangeul_static/css/nanum-square-neo.css" rel="stylesheet"><style>body{font-family:NanumSquareNeo,sans-serif;font-size:' + fontSize + 'px;}</style></head><body>' + bodyParts.join('') + '</body></html>';
        saveBlob(new Blob(['\ufeff' + html], { type: 'application/msword' }), getFilename('.doc'));
        return;
      }
    }

    if (type === 'text') {
      var textLines;
      if (hasFull) {
        textLines = fullRows.map(function (r) {
          return '단원: ' + r.unitNo + ' | ' + r.unitName + ' | 문항: ' + r.questionNo +
            '\n[지시문] ' + (r.instruction || '') +
            '\n[지문] ' + (r.passage || '') +
            '\n[선택지] ' + (r.choices || '') +
            (r.words ? '\n[단어] ' + r.words : '');
        });
      } else {
        textLines = extractedData.map(function (r) {
          return '단원: ' + r.unitNo + (r.unitName ? ' | ' + r.unitName : '') + ' | 문항: ' + r.questionNo + '\n' + r.content;
        });
      }
      const blob = new Blob([textLines.join('\n\n')], { type: 'text/plain;charset=utf-8' });
      saveBlob(blob, getFilename('.txt'));
      return;
    }

    if (type === 'excel') {
      var wsData;
      if (hasFull) {
        wsData = [['단원번호', '단원명', '문항번호', '지시문', '지문', '선택지', '단어']];
        fullRows.forEach(function (r) {
          wsData.push([r.unitNo, r.unitName, r.questionNo, r.instruction, r.passage, r.choices, r.words]);
        });
      } else {
        wsData = [['단원번호', '단원명', '문항번호', '내용']];
        extractedData.forEach(function (r) {
          wsData.push([r.unitNo, r.unitName, r.questionNo, r.content]);
        });
      }
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(wsData);
      XLSX.utils.book_append_sheet(wb, ws, hasFull ? '문항데이터' : '스크립트');
      XLSX.writeFile(wb, getFilename('.xlsx'));
      return;
    }

    if (type === 'word') {
      var rows, headers;
      if (hasFull) {
        headers = '<tr><th>단원번호</th><th>단원명</th><th>문항번호</th><th>지시문</th><th>지문</th><th>선택지</th><th>단어</th></tr>';
        rows = fullRows.map(function (r) {
          return '<tr><td>' + escapeHtml(r.unitNo) + '</td><td>' + escapeHtml(r.unitName) + '</td><td>' + escapeHtml(r.questionNo) + '</td><td>' + escapeHtml(r.instruction) + '</td><td>' + escapeHtml(r.passage) + '</td><td>' + escapeHtml(r.choices) + '</td><td>' + escapeHtml(r.words) + '</td></tr>';
        }).join('');
      } else {
        headers = '<tr><th>단원번호</th><th>단원명</th><th>문항번호</th><th>내용</th></tr>';
        rows = extractedData.map(function (r) {
          return '<tr><td>' + escapeHtml(r.unitNo) + '</td><td>' + escapeHtml(r.unitName) + '</td><td>' + escapeHtml(r.questionNo) + '</td><td>' + escapeHtml(r.content) + '</td></tr>';
        }).join('');
      }
      const html = '<!DOCTYPE html><html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word"><head><meta charset="UTF-8"><link href="https://hangeul.pstatic.net/hangeul_static/css/nanum-square-neo.css" rel="stylesheet"><style>body{font-family:NanumSquareNeo,sans-serif;font-size:' + fontSize + 'px;} table{border-collapse:collapse;width:100%;} th,td{border:1px solid #ddd;padding:8px;}</style></head><body><table><thead>' + headers + '</thead><tbody>' + rows + '</tbody></table></body></html>';
      const blob = new Blob(['\ufeff' + html], { type: 'application/msword' });
      saveBlob(blob, getFilename('.doc'));
    }
  }

  function saveBlob(blob, filename) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
  }
})();
