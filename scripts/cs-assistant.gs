// CS 답변 초안 (admin.html "CS 상담" 탭) — Apps Script 백엔드
//
// 설치 방법
// 1. 주문 스프레드시트의 Apps Script 편집기에서 파일 추가(+) > 스크립트 > 이름 "cs-assistant"
//    로 만들고 이 파일 내용을 그대로 붙여넣습니다.
// 2. 기존 doPost(e) 안, `var data = JSON.parse(e.postData.contents);` 바로 아래에
//    아래 한 줄을 추가합니다.
//        if (String(data.action || '').indexOf('cs') === 0) return handleCsAction_(data);
// 3. 프로젝트 설정 > 스크립트 속성에 ANTHROPIC_API_KEY (sk-ant-...) 를 추가합니다.
//    선택: CS_MODEL (기본 claude-opus-5-5), CS_EFFORT (기본 high — low/medium/high/xhigh/max)
// 4. 배포 > 배포 관리 > 기존 웹앱 편집 > 버전 "새 버전" 으로 재배포합니다. (URL 은 그대로 유지)
// 5. admin.html > CS 상담 > "상담 지침" 에 Claude 프로젝트 지침(CLAUDE.md) 내용을 붙여넣고 저장합니다.
//
// 지침은 공개 저장소에 올리지 않고 스프레드시트의 숨김 시트("CS지침")에만 저장됩니다.
// API 키는 스크립트 속성에만 있고 브라우저로 전달되지 않습니다.

var CS_GUIDE_SHEET = 'CS지침';
var CS_CHUNK = 40000; // 셀 하나 최대 50,000자 — 여유 있게 나눠 저장
var CS_DEFAULT_MODEL = 'claude-opus-5-5';
var CS_DEFAULT_EFFORT = 'high';

var CS_SYSTEM_PREAMBLE = [
  '너는 쿠팡에서 신선 농산물을 판매하는 셀러의 CS 상담 담당자다.',
  '관리자가 고객 문의(텍스트 또는 쿠팡 Wing 화면 캡처)를 주면, 아래 <cs_guide> 지침을 그대로 따라 바로 등록 가능한 답변 초안을 쓴다.',
  '',
  '작업 방식',
  '- 캡처 이미지가 있으면 상품명, 문의 날짜, 고객 문의 원문, 이전 판매자 답변, 고객이 올린 사진 유무를 먼저 꼼꼼히 읽어낸다. 여러 장이면 시간 순서대로 이어서 읽는다.',
  '- 지침의 "판단 순서"대로 위험 신호 → 재문의 여부 → 상황별 기준 표 → 템플릿 순으로 판단한다.',
  '- 템플릿은 뼈대로 쓰고, 고객이 실제로 쓴 표현과 상황을 살려 자연스럽게 다시 쓴다. 템플릿 문장을 기계적으로 붙여넣지 않는다.',
  '- 출력은 지침의 "출력 형식"만 지킨다. 형식 밖의 설명, 머리말, 코드블록 표시는 쓰지 않는다.',
  '- 관리자가 수정을 요청하면 같은 형식으로 전체 답변을 다시 쓴다.',
  '- 지침에 없는 사실(발송 내역, 산지, 보상 결정 등)은 지어내지 않고 "확인 필요"로 표시한다.'
].join('\n');

function handleCsAction_(data) {
  try {
    if (!isValidAdminKey_(data.adminKey)) return jsonOutput_({ ok: false, error: '인증에 실패했습니다.' });
    if (data.action === 'csGetGuide') return jsonOutput_(csGetGuide_());
    if (data.action === 'csSaveGuide') return jsonOutput_(csSaveGuide_(data.guide));
    if (data.action === 'csDraft') return jsonOutput_(csDraft_(data.messages));
    return jsonOutput_({ ok: false, error: '알 수 없는 요청입니다.' });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err && err.message || err) });
  }
}

function csGuideSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CS_GUIDE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CS_GUIDE_SHEET);
    sheet.hideSheet();
  }
  return sheet;
}

function csReadGuide_() {
  var sheet = csGuideSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return '';
  return sheet.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); }).join('');
}

function csGetGuide_() {
  var sheet = csGuideSheet_();
  var props = PropertiesService.getScriptProperties();
  return {
    ok: true,
    guide: csReadGuide_(),
    updatedAt: String(sheet.getRange(1, 1).getValue() || ''),
    hasApiKey: !!props.getProperty('ANTHROPIC_API_KEY'),
    model: props.getProperty('CS_MODEL') || CS_DEFAULT_MODEL
  };
}

function csSaveGuide_(guide) {
  guide = String(guide || '').trim();
  if (!guide) return { ok: false, error: '지침 내용이 비어 있습니다.' };
  var sheet = csGuideSheet_();
  sheet.clearContents();
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var rows = [[stamp]];
  for (var i = 0; i < guide.length; i += CS_CHUNK) rows.push([guide.slice(i, i + CS_CHUNK)]);
  sheet.getRange(1, 1, rows.length, 1).setNumberFormat('@').setValues(rows);
  return { ok: true, updatedAt: stamp };
}

function csDraft_(messages) {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { ok: false, error: '스크립트 속성에 ANTHROPIC_API_KEY 가 없습니다.' };
  var guide = csReadGuide_();
  if (!guide) return { ok: false, error: '상담 지침이 저장되어 있지 않습니다. CS 상담 > 상담 지침에 먼저 붙여넣어 주세요.' };
  if (!Array.isArray(messages) || !messages.length) return { ok: false, error: '문의 내용이 없습니다.' };
  messages = messages.filter(function (m) {
    return m && (m.role === 'user' || m.role === 'assistant') && m.content;
  });
  if (!messages.length || messages[messages.length - 1].role !== 'user') return { ok: false, error: '요청 형식이 올바르지 않습니다.' };

  var body = {
    model: props.getProperty('CS_MODEL') || CS_DEFAULT_MODEL,
    max_tokens: 16000,
    output_config: { effort: props.getProperty('CS_EFFORT') || CS_DEFAULT_EFFORT },
    // 안전 분류기가 드물게 거절하면 Anthropic 권장 모델로 자동 재시도
    fallbacks: 'default',
    system: [{
      type: 'text',
      text: CS_SYSTEM_PREAMBLE + '\n\n<cs_guide>\n' + guide + '\n</cs_guide>',
      cache_control: { type: 'ephemeral' } // 지침은 매번 같으므로 캐시해서 비용·속도 절약
    }],
    messages: messages
  };

  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2026-10-21',
      'anthropic-beta': 'server-side-fallback-2026-07-01'
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  var code = res.getResponseCode();
  var json;
  try { json = JSON.parse(res.getContentText()); } catch (e) { json = null; }
  if (code !== 200 || !json) {
    var msg = json && json.error && json.error.message ? json.error.message : res.getContentText().slice(0, 300);
    if (code === 429 || code === 529) msg = 'AI 서버가 바쁩니다. 잠시 후 다시 시도해주세요. (' + msg + ')';
    return { ok: false, error: 'AI 요청 실패 (' + code + '): ' + msg };
  }
  if (json.stop_reason === 'refusal') {
    return { ok: false, error: 'AI가 이 요청에 답변하지 않았습니다. 문의 내용을 조금 바꿔 다시 시도해주세요.' };
  }

  var text = (json.content || [])
    .filter(function (b) { return b.type === 'text'; })
    .map(function (b) { return b.text; })
    .join('')
    .trim();

  return {
    ok: true,
    text: text,
    content: json.content, // 수정 요청 시 대화를 이어가기 위해 그대로 돌려줌
    truncated: json.stop_reason === 'max_tokens',
    usage: json.usage
  };
}
