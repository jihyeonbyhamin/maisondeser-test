// CS 상담 탭 — 쿠팡 문의 답변 초안 (백엔드: scripts/cs-assistant.gs)
(function () {
  'use strict';
  var MAX_IMAGES = 20;
  var cs = { images: [], messages: [], lastText: '', lastBody: '', guideLoaded: false, busy: false };
  function $(s) { return document.querySelector(s); }
  function A() { return window.mdsAdmin; }
  function toast(m, e) { A().toast(m, e); }

  // ---------- 지침 ----------
  function loadGuide() {
    $('#csStatus').textContent = '지침 확인 중…';
    return A().api({ action: 'csGetGuide' }).then(function (r) {
      if (!r.ok) throw new Error(r.error || '불러오기 실패');
      cs.guideLoaded = true;
      $('#csGuide').value = r.guide || '';
      $('#csGuideMeta').textContent = r.guide ? '· 마지막 저장 ' + (r.updatedAt || '-') : '· 아직 저장되지 않음';
      if (!r.guide) $('#csGuideBox').open = true;
      setStatus(r);
    }).catch(function (e) {
      $('#csStatus').textContent = 'Apps Script 재배포 필요';
      $('#csStatus').className = 'cs-status warn';
      $('#csGuideNote').textContent = 'CS 백엔드(cs-assistant.gs)를 설치·재배포한 뒤 새로고침하세요. (' + e.message + ')';
    });
  }
  function setStatus(r) {
    var el = $('#csStatus');
    if (!r.hasApiKey) { el.textContent = 'API 키 설정 필요'; el.className = 'cs-status warn'; return; }
    if (!r.guide) { el.textContent = '상담 지침 저장 필요'; el.className = 'cs-status warn'; return; }
    el.textContent = 'AI 연결됨 · ' + (r.model || '');
    el.className = 'cs-status ok';
  }
  function saveGuide() {
    var guide = $('#csGuide').value.trim();
    if (!guide) { toast('지침 내용을 붙여넣어 주세요.', true); return; }
    $('#csGuideNote').textContent = '저장 중…';
    A().api({ action: 'csSaveGuide', guide: guide }).then(function (r) {
      if (!r.ok) throw new Error(r.error);
      $('#csGuideNote').textContent = '저장되었습니다.';
      $('#csGuideMeta').textContent = '· 마지막 저장 ' + r.updatedAt;
      toast('상담 지침을 저장했습니다.');
      loadGuide();
    }).catch(function (e) { $('#csGuideNote').textContent = e.message || '저장 실패'; toast('지침 저장에 실패했습니다.', true); });
  }

  // ---------- 이미지 ----------
  // 긴 캡처는 글자가 뭉개지지 않도록 세로로 나눠 보냄
  function processImage(file) {
    return new Promise(function (resolve, reject) {
      var url = URL.createObjectURL(file), img = new Image();
      img.onload = function () {
        URL.revokeObjectURL(url);
        var scale = Math.min(1, 1400 / img.width);
        var w = Math.round(img.width * scale), h = Math.round(img.height * scale);
        var sliceH = Math.round(w * 1.6), overlap = 60, parts = [];
        if (h <= w * 2) sliceH = h;
        for (var y = 0; y < h; y += sliceH - (sliceH < h ? overlap : 0)) {
          var ph = Math.min(sliceH, h - y), c = document.createElement('canvas');
          c.width = w; c.height = ph;
          var ctx = c.getContext('2d');
          ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, w, ph);
          ctx.drawImage(img, 0, y / scale, img.width, ph / scale, 0, 0, w, ph);
          parts.push(c.toDataURL('image/jpeg', 0.9));
          if (y + ph >= h) break;
        }
        resolve(parts);
      };
      img.onerror = function () { URL.revokeObjectURL(url); reject(new Error('이미지를 읽을 수 없습니다.')); };
      img.src = url;
    });
  }
  function addFiles(files) {
    files = Array.prototype.filter.call(files || [], function (f) { return /^image\//.test(f.type); });
    if (!files.length) return;
    Promise.all(files.map(processImage)).then(function (groups) {
      groups.forEach(function (parts) { cs.images.push(parts); });
      var total = cs.images.reduce(function (n, g) { return n + g.length; }, 0);
      if (total > MAX_IMAGES) toast('이미지가 많아 앞쪽 ' + MAX_IMAGES + '조각까지만 보냅니다.', true);
      renderThumbs();
    }).catch(function (e) { toast(e.message, true); });
  }
  function renderThumbs() {
    var wrap = $('#csThumbs'); wrap.innerHTML = '';
    cs.images.forEach(function (parts, i) {
      var d = document.createElement('div'); d.className = 'cs-thumb';
      d.innerHTML = '<img alt=""><button type="button" aria-label="삭제">×</button>' + (parts.length > 1 ? '<em>' + parts.length + '조각</em>' : '');
      d.querySelector('img').src = parts[0];
      d.querySelector('button').onclick = function () { cs.images.splice(i, 1); renderThumbs(); };
      wrap.appendChild(d);
    });
    $('#csDrop').classList.toggle('has-images', cs.images.length > 0);
  }
  function imageBlocks() {
    var flat = [].concat.apply([], cs.images).slice(0, MAX_IMAGES);
    return flat.map(function (dataUrl) {
      return { type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: dataUrl.split(',')[1] } };
    });
  }

  // ---------- 요청 ----------
  function today() { var d = new Date(); return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
  function buildFirstMessage() {
    var v = function (id) { return $(id).value.trim(); };
    var product = v('#csProduct'), inquiry = v('#csInquiry'), history = v('#csHistory'), stage = v('#csStage'), photos = v('#csPhotos'), memo = v('#csMemo');
    var images = imageBlocks();
    if (!inquiry && !images.length) return null;
    var lines = ['오늘 날짜: ' + today()];
    if (images.length) lines.push('[캡처] 쿠팡 문의 화면 캡처 ' + images.length + '장을 첨부했습니다. 위에서부터 순서대로 이어진 화면입니다.');
    lines.push('[상품명] ' + (product || (images.length ? '캡처에서 확인' : '입력 없음')));
    if (inquiry) lines.push('[고객 문의]\n' + inquiry);
    lines.push('[이전 판매자 답변 / 대화 이력]\n' + (history || (images.length ? '별도 입력 없음 — 캡처에 있으면 그것을 따르세요.' : '입력 없음')));
    if (stage) lines.push('[문의 차수] ' + stage);
    if (photos) lines.push('[고객 사진] ' + photos);
    if (memo) lines.push('[대표님 지시 / 메모 — 지침보다 우선]\n' + memo);
    lines.push('위 문의에 대한 답변 초안을 지침의 출력 형식대로 작성해주세요.');
    return { role: 'user', content: images.concat([{ type: 'text', text: lines.join('\n\n') }]) };
  }
  function send(messages) {
    if (cs.busy) return;
    cs.busy = true;
    $('#csSubmit').disabled = true;
    $('#csEmpty').hidden = true; $('#csLoading').hidden = false;
    $('#csResult').classList.add('dim');
    A().api({ action: 'csDraft', messages: messages }).then(function (r) {
      if (!r.ok) throw new Error(r.error || '답변 생성 실패');
      cs.messages = messages.concat([{ role: 'assistant', content: r.content }]);
      showResult(r.text);
      if (r.truncated) toast('답변이 길어 중간에 끊겼습니다. "이어서 써줘"로 요청해주세요.', true);
    }).catch(function (e) {
      toast(e.message || '답변 생성에 실패했습니다.', true);
      if (!cs.lastText) $('#csEmpty').hidden = false;
    }).finally(function () {
      cs.busy = false; $('#csSubmit').disabled = false; $('#csLoading').hidden = true;
      $('#csResult').classList.remove('dim');
    });
  }
  function generate(e) {
    if (e) e.preventDefault();
    if (!A().isLoggedIn()) return;
    var first = buildFirstMessage();
    if (!first) { toast('문의 내용을 입력하거나 캡처를 넣어주세요.', true); return; }
    send([first]);
  }
  function refine(e) {
    e.preventDefault();
    var req = $('#csRefineInput').value.trim();
    if (!req || !cs.messages.length) return;
    var text = req;
    var edited = $('#csAnswer').value.trim();
    if (edited && edited !== cs.lastBody) text = '(관리자가 직접 고친 현재 본문)\n' + edited + '\n\n요청: ' + req;
    $('#csRefineInput').value = '';
    send(cs.messages.concat([{ role: 'user', content: text }]));
  }

  // ---------- 결과 ----------
  function splitAnswer(text) {
    text = text.replace(/^```[a-z]*\n?|\n?```$/g, '').trim();
    var idx = text.indexOf('\n\n'), head = idx > -1 ? text.slice(0, idx) : '';
    if (head && /^\[|\/|확인 필요|⚠️/.test(head) && !/^안녕하세요/.test(head)) return { head: head.split('\n'), body: text.slice(idx + 2).trim() };
    return { head: [], body: text };
  }
  function showResult(text) {
    cs.lastText = text;
    var parts = splitAnswer(text);
    cs.lastBody = parts.body;
    var meta = $('#csMeta'); meta.innerHTML = '';
    parts.head.forEach(function (line) {
      var p = document.createElement('p');
      p.textContent = line;
      if (/⚠️/.test(line) && !/주의 없음|⚠️ 없음|주의 여부: 없음/.test(line)) p.className = 'warn';
      if (/^확인 필요/.test(line)) p.className = 'check';
      meta.appendChild(p);
    });
    meta.hidden = !parts.head.length;
    $('#csAnswer').value = parts.body;
    updateCount();
    $('#csResult').hidden = false;
    $('#csEmpty').hidden = true;
  }
  function updateCount() { $('#csCount').textContent = $('#csAnswer').value.length.toLocaleString('ko-KR') + '자'; }
  function copy(text, label) {
    (navigator.clipboard ? navigator.clipboard.writeText(text) : Promise.reject()).then(function () { toast(label + ' 복사했습니다.'); })
      .catch(function () { var t = $('#csAnswer'); t.focus(); t.select(); document.execCommand('copy'); toast(label + ' 복사했습니다.'); });
  }
  function reset() {
    ['#csProduct', '#csInquiry', '#csHistory', '#csMemo', '#csRefineInput'].forEach(function (id) { $(id).value = ''; });
    $('#csStage').value = ''; $('#csPhotos').value = '';
    cs.images = []; cs.messages = []; cs.lastText = ''; cs.lastBody = '';
    renderThumbs();
    $('#csResult').hidden = true; $('#csEmpty').hidden = false;
    $('#csInquiry').focus();
  }

  // ---------- 이벤트 ----------
  document.querySelector('.admin-tabs').addEventListener('click', function (e) {
    var b = e.target.closest('[data-tab]');
    if (b && b.dataset.tab === 'cs' && !cs.guideLoaded) loadGuide();
  });
  document.addEventListener('paste', function (e) {
    var pane = document.querySelector('.admin-pane[data-pane="cs"]');
    if (!pane.classList.contains('active') || e.target.id === 'csGuide') return;
    var files = Array.prototype.map.call((e.clipboardData && e.clipboardData.items) || [], function (it) { return it.kind === 'file' ? it.getAsFile() : null; }).filter(Boolean);
    if (files.length) { e.preventDefault(); addFiles(files); }
  });
  var drop = $('#csDrop');
  ['dragenter', 'dragover'].forEach(function (t) { drop.addEventListener(t, function (e) { e.preventDefault(); drop.classList.add('over'); }); });
  ['dragleave', 'drop'].forEach(function (t) { drop.addEventListener(t, function (e) { e.preventDefault(); drop.classList.remove('over'); }); });
  drop.addEventListener('drop', function (e) { addFiles(e.dataTransfer.files); });
  $('#csFile').addEventListener('change', function () { addFiles(this.files); this.value = ''; });
  $('#csForm').addEventListener('submit', generate);
  $('#csForm').addEventListener('keydown', function (e) { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) generate(e); });
  $('#csRefine').addEventListener('submit', refine);
  $('#csReset').onclick = reset;
  $('#csAnswer').addEventListener('input', updateCount);
  $('#csCopy').onclick = function () { copy($('#csAnswer').value, '답변 본문을'); };
  $('#csCopyAll').onclick = function () {
    var head = Array.prototype.map.call(document.querySelectorAll('#csMeta p'), function (p) { return p.textContent; }).join('\n');
    copy((head ? head + '\n\n' : '') + $('#csAnswer').value, '유형과 본문을');
  };
  $('#csGuideSave').onclick = saveGuide;
})();
