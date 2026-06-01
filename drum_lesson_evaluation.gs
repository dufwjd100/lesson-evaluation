/**
 * 드럼 레슨평가 초안 자동 생성 스크립트
 * 수원점 출석부 기반 / 학부모용 / 수동 실행
 */

// ───── 설정 ─────
const ATTENDANCE_SHEET_ID = '1QHrIr76dBf49rDWrN2I5IVaZgAZSi8tAScNRW0SRxt4';
const CURRICULUM_SHEET_ID = '18omuRClflSU5JkQLlbXmfiqQyu8tOnvvG2TP8KWFMwA';
const ATTENDANCE_TAB      = '출석부';
const CURRICULUM_TAB      = '드럼_커리큘럼_V4';

// OpenAI API 키 (스크립트 속성에서 가져옴)
function getApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!key) throw new Error('스크립트 속성에 OPENAI_API_KEY 가 설정되어 있지 않습니다.');
  return key;
}

// ───── 메뉴 등록 ─────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🥁 레슨평가')
    .addItem('드럼 레슨평가 초안 생성', 'generateDrumEvaluations')
    .addToUi();
}

// ───── 메인 ─────
function generateDrumEvaluations() {
  const ui = SpreadsheetApp.getUi();

  try {
    const attendanceSS  = SpreadsheetApp.openById(ATTENDANCE_SHEET_ID);
    const attendanceSheet = attendanceSS.getSheetByName(ATTENDANCE_TAB);
    if (!attendanceSheet) throw new Error(`"${ATTENDANCE_TAB}" 탭을 찾을 수 없습니다.`);

    const curriculumSS    = SpreadsheetApp.openById(CURRICULUM_SHEET_ID);
    const curriculumSheet = curriculumSS.getSheetByName(CURRICULUM_TAB);
    if (!curriculumSheet) throw new Error(`"${CURRICULUM_TAB}" 탭을 찾을 수 없습니다.`);

    // 1. 헤더 파싱
    const headers = attendanceSheet.getRange(1, 1, 1, attendanceSheet.getLastColumn())
      .getValues()[0]
      .map(h => String(h).trim());

    const COL = buildColIndex(headers);

    // 2. 전체 데이터 읽기 (2행부터)
    const lastRow  = attendanceSheet.getLastRow();
    if (lastRow < 2) { ui.alert('데이터가 없습니다.'); return; }

    const dataRange = attendanceSheet.getRange(2, 1, lastRow - 1, headers.length);
    const allData   = dataRange.getValues();

    // 3. 커리큘럼 읽기
    const curriculum = loadCurriculum(curriculumSheet);

    // 4. 대상 행 필터링
    const targets = findTargetRows(allData, COL);

    if (targets.length === 0) {
      ui.alert('처리할 대상 행이 없습니다.\n(드럼 수업, 진행회차 4의 배수, 레슨평가_초안 비어있음, 고유번호 있음)');
      return;
    }

    const confirm = ui.alert(
      `총 ${targets.length}건 처리합니다. 계속하시겠습니까?`,
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;

    let successCount = 0;
    let failCount    = 0;
    const errors     = [];

    for (const targetIdx of targets) {
      try {
        const targetRow = allData[targetIdx];
        const uid       = String(targetRow[COL.고유번호]).trim();

        // 5. 같은 고유번호의 최근 4회 수업기록 수집
        const recentLessons = getRecentLessons(allData, COL, uid, targetIdx);

        // 6. 커리큘럼 매칭
        const matchedSteps = matchCurriculum(recentLessons, curriculum);

        // 7. AI 초안 생성
        const draft = callOpenAIAPI(recentLessons, matchedSteps, curriculum);

        // 8. 출석부에 저장 (row index → sheet row: +2 because 1-indexed and header)
        const sheetRow = targetIdx + 2;
        attendanceSheet.getRange(sheetRow, COL.레슨평가_초안 + 1).setValue(draft);
        SpreadsheetApp.flush();

        successCount++;
        Utilities.sleep(1500); // API 과부하 방지
      } catch (e) {
        failCount++;
        errors.push(`행 ${targetIdx + 2}: ${e.message}`);
      }
    }

    let msg = `완료: ${successCount}건 성공`;
    if (failCount > 0) msg += `, ${failCount}건 실패\n\n` + errors.join('\n');
    ui.alert(msg);

  } catch (e) {
    ui.alert('오류: ' + e.message);
    console.error(e);
  }
}

// ───── 헬퍼: 컬럼 인덱스 맵 ─────
function buildColIndex(headers) {
  const map = {};
  headers.forEach((h, i) => { map[h] = i; });

  const required = ['고유번호', '수업명', '진행회차', '수업일자', '수업내용', '수업곡', '과제', '레슨평가', '레슨평가_초안', '출결'];
  for (const col of required) {
    if (map[col] === undefined) {
      throw new Error(`헤더에서 "${col}" 컬럼을 찾을 수 없습니다. 실제 헤더: ${headers.join(', ')}`);
    }
  }
  return map;
}

// ───── 헬퍼: 대상 행 탐색 ─────
function findTargetRows(allData, COL) {
  const targets = [];
  allData.forEach((row, i) => {
    const lessonName   = String(row[COL.수업명] || '').trim();
    const seq          = Number(row[COL.진행회차]);
    const draftExists  = String(row[COL.레슨평가_초안] || '').trim();
    const uid          = String(row[COL.고유번호] || '').trim();

    if (
      lessonName.includes('드럼') &&
      seq > 0 &&
      seq % 4 === 0 &&
      draftExists === '' &&
      uid !== ''
    ) {
      targets.push(i);
    }
  });
  return targets;
}

// ───── 헬퍼: 최근 4회 수업기록 수집 ─────
function getRecentLessons(allData, COL, uid, upToIdx) {
  // 같은 고유번호, upToIdx 이하, 결석 제외, 날짜 내림차순으로 최대 4개
  const lessons = [];
  for (let i = upToIdx; i >= 0; i--) {
    const row = allData[i];
    if (String(row[COL.고유번호]).trim() !== uid) continue;

    const attendance = String(row[COL.출결] || '').trim();
    if (['결석', '무단결석'].includes(attendance)) continue;

    lessons.push({
      date    : row[COL.수업일자],
      seq     : row[COL.진행회차],
      content : String(row[COL.수업내용] || '').trim(),
      song    : String(row[COL.수업곡]   || '').trim(),
      homework: String(row[COL.과제]     || '').trim(),
      eval    : String(row[COL.레슨평가] || '').trim(),
    });

    if (lessons.length >= 4) break;
  }
  return lessons.reverse(); // 오래된 순으로
}

// ───── 커리큘럼 로드 ─────
function loadCurriculum(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const data    = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const COL = {};
  headers.forEach((h, i) => { COL[h] = i; });

  return data.map(row => ({
    step          : String(row[COL['Step']]              || '').trim(),
    coreGoal      : String(row[COL['핵심목표']]           || '').trim(),
    basicEasy     : String(row[COL['기본기 쉬운해석']]    || '').trim(),
    sequenceEasy  : String(row[COL['시퀀스 쉬운해석']]   || '').trim(),
    evalSentence  : String(row[COL['평가서 조합 문장']]  || '').trim(),
    futureSentence: String(row[COL['미래지향 문장']]      || '').trim(),
    keywords      : String(row[COL['자동화 검색 키워드']] || '').trim(),
  })).filter(r => r.step !== '');
}

// ───── 커리큘럼 매칭 ─────
function matchCurriculum(lessons, curriculum) {
  const allContent = lessons.map(l => l.content).join(' ');

  // 각 Step의 키워드 매칭 점수 계산
  const scores = curriculum.map(step => {
    if (!step.keywords) return { step, score: 0 };
    const kws    = step.keywords.split(/[,，\s]+/).map(k => k.trim()).filter(Boolean);
    const score  = kws.filter(kw => allContent.includes(kw)).length;
    return { step, score };
  }).filter(s => s.score > 0);

  scores.sort((a, b) => b.score - a.score);

  // 1위 + 보조 최대 2개
  return scores.slice(0, 3).map(s => s.step);
}

// ───── OpenAI API 호출 ─────
function callOpenAIAPI(lessons, matchedSteps, curriculum) {
  const apiKey = getApiKey();

  const lessonSummary = lessons.map((l, i) =>
    `[${i + 1}회] 날짜:${l.date} 수업내용:${l.content} 수업곡:${l.song} 과제:${l.homework} 기존평가:${l.eval}`
  ).join('\n');

  const stepSummary = matchedSteps.map(s =>
    `Step ${s.step}\n핵심목표: ${s.coreGoal}\n기본기: ${s.basicEasy}\n시퀀스: ${s.sequenceEasy}\n평가문장: ${s.evalSentence}\n미래방향: ${s.futureSentence}`
  ).join('\n\n');

  const prompt = `당신은 음악학원의 드럼 강사입니다. 아래 최근 4회 수업기록과 커리큘럼 정보를 바탕으로 학부모용 레슨평가 초안을 작성해 주세요.

【최근 수업기록】
${lessonSummary}

【커리큘럼 참고】
${stepSummary || '(매칭된 커리큘럼 없음)'}

【작성 규칙】
- 3문단으로 구성
  1문단: 최근 4회 동안 무엇을 배웠는지 (수업내용 근거 기반)
  2문단: 좋아진 점과 현재 상태
  3문단: 앞으로의 수업 방향
- 학부모가 읽는 글이므로 전문용어를 쉬운 말로 풀어씀
- 따뜻하지만 과장하지 않음 ("정말 대단해요" 식의 과도한 칭찬 금지)
- "잘하고 있습니다"만 반복하지 않음
- 수업내용에 근거하지 않은 내용 추가 금지
- 분량: 200~300자 내외

평가 초안만 출력하세요. 설명이나 제목 없이 본문만 작성합니다.`;

  const payload = {
    model: 'gpt-4.1-mini',
    input: prompt,
    max_output_tokens: 600,
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${apiKey}`,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch('https://api.openai.com/v1/responses', options);
  const code = response.getResponseCode();
  const rawText = response.getContentText();
  const body = JSON.parse(rawText);

  if (code < 200 || code >= 300) {
    throw new Error(`API 오류 (${code}): ${body.error?.message || rawText}`);
  }

  const outputText = extractOpenAIText(body);
  if (!outputText) {
    throw new Error('API 응답에서 평가 초안 텍스트를 찾을 수 없습니다.');
  }

  return outputText.trim();
}

// ───── OpenAI 응답 텍스트 추출 ─────
function extractOpenAIText(body) {
  if (body.output_text) return String(body.output_text).trim();

  if (!body.output || !Array.isArray(body.output)) return '';

  const parts = [];
  body.output.forEach(item => {
    if (!item.content || !Array.isArray(item.content)) return;
    item.content.forEach(content => {
      if (content.text) parts.push(String(content.text));
    });
  });

  return parts.join('\n').trim();
}
