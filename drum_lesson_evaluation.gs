/**
 * 드럼 레슨평가 초안 자동 생성 스크립트
 * 수원점 출석부 기반 / 학부모용 / 수동 실행
 */

// ───── 설정 ─────
const ATTENDANCE_SHEET_ID = '1QHrIr76dBf49rDWrN2I5IVaZgAZSi8tAScNRW0SRxt4';
const CURRICULUM_SHEET_ID = '18omuRClflSU5JkQLlbXmfiqQyu8tOnvvG2TP8KWFMwA';
const ATTENDANCE_TAB      = '출석부';
const CURRICULUM_TAB      = '드럼_커리큘럼_V4';

// Anthropic API 키 (스크립트 속성에서 가져옴)
function getApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) throw new Error('스크립트 속성에 ANTHROPIC_API_KEY 가 설정되어 있지 않습니다.');
  return key;
}

// ───── 메뉴 등록 ─────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🥁 레슨평가')
    .addItem('선택한 행만 레슨평가 초안 생성', 'generateSelectedDrumEvaluation')
    .addSeparator()
    .addItem('전체 대상 레슨평가 초안 생성', 'generateDrumEvaluations')
    .addToUi();
}

// ───── 선택한 행 1건만 처리 ─────
function generateSelectedDrumEvaluation() {
  const ui = SpreadsheetApp.getUi();

  try {
    const activeSheet = SpreadsheetApp.getActiveSheet();
    if (!activeSheet || activeSheet.getName() !== ATTENDANCE_TAB) {
      throw new Error(`"${ATTENDANCE_TAB}" 탭에서 처리할 행을 선택한 뒤 실행해 주세요.`);
    }

    const selectedRow = activeSheet.getActiveRange().getRow();
    if (selectedRow <= 1) {
      throw new Error('헤더가 아닌 실제 수업 기록 행을 선택해 주세요.');
    }

    const headers = activeSheet.getRange(1, 1, 1, activeSheet.getLastColumn())
      .getValues()[0]
      .map(h => String(h).trim());
    const COL = buildColIndex(headers);

    const lastRow = activeSheet.getLastRow();
    const allData = activeSheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    const targetIdx = selectedRow - 2;
    const targetRow = allData[targetIdx];

    if (!targetRow) throw new Error('선택한 행의 데이터를 찾을 수 없습니다.');

    const uid = String(targetRow[COL.고유번호] || '').trim();
    const lessonName = String(targetRow[COL.수업명] || '').trim();
    const currentDraft = String(targetRow[COL.레슨평가_초안] || '').trim();

    if (!uid) throw new Error('선택한 행에 고유번호가 없습니다.');
    if (!lessonName.includes('드럼')) throw new Error('선택한 행의 수업명이 드럼 수업이 아닙니다.');

    if (currentDraft) {
      const overwrite = ui.alert(
        '이미 레슨평가_초안이 있습니다. 덮어쓸까요?',
        ui.ButtonSet.YES_NO
      );
      if (overwrite !== ui.Button.YES) return;
    }

    const curriculumSS = SpreadsheetApp.openById(CURRICULUM_SHEET_ID);
    const curriculumSheet = curriculumSS.getSheetByName(CURRICULUM_TAB);
    if (!curriculumSheet) throw new Error(`"${CURRICULUM_TAB}" 탭을 찾을 수 없습니다.`);

    const curriculum = loadCurriculum(curriculumSheet);
    const recentLessons = getRecentLessons(allData, COL, uid, targetIdx);
    if (recentLessons.length === 0) {
      throw new Error('선택한 학생의 최근 수업기록을 찾을 수 없습니다.');
    }

    const matchedSteps = matchCurriculum(recentLessons, curriculum);
    const draft = callOpenAIAPI(recentLessons, matchedSteps, curriculum);

    activeSheet.getRange(selectedRow, COL.레슨평가_초안 + 1).setValue(draft);
    SpreadsheetApp.flush();

    ui.alert(`선택한 행 레슨평가 초안 생성 완료\n행: ${selectedRow}\n고유번호: ${uid}`);

  } catch (e) {
    ui.alert('오류: ' + e.message);
    console.error(e);
  }
}

// ───── 전체 대상 처리 ─────
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

  // 내부 코드명은 유지하고, 실제 시트 헤더명만 별칭으로 연결한다.
  const aliases = {
    수업일자: ['수업일자', '날짜'],
    출결: ['출결', '출결상황'],
  };

  Object.keys(aliases).forEach(key => {
    if (map[key] !== undefined) return;
    const found = aliases[key].find(name => map[name] !== undefined);
    if (found) map[key] = map[found];
  });

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
      name    : COL.이름 !== undefined ? String(row[COL.이름] || '').trim() : '',
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
    section          : getCellValue(row, COL, ['구간']),
    step             : getCellValue(row, COL, ['Step']),
    stepName         : getCellValue(row, COL, ['단계명']),
    stepRole         : getCellValue(row, COL, ['Step 역할']),
    coreGoal         : getCellValue(row, COL, ['핵심목표']),
    basicArea        : getCellValue(row, COL, ['기본기 세부영역']),
    relatedDrill     : getCellValue(row, COL, ['관련 Drill']),
    basicEasy        : getCellValue(row, COL, ['기본기 쉬운해석']),
    sequenceArea     : getCellValue(row, COL, ['시퀀스 세부영역']),
    relatedPattern   : getCellValue(row, COL, ['관련 Cell/Groove/Fill']),
    sequenceEasy     : getCellValue(row, COL, ['시퀀스 쉬운해석']),
    songApply        : getCellValue(row, COL, ['곡 적용 기준']),
    tempoHint        : getCellValue(row, COL, ['템포·속도 힌트']),
    contextKeywords  : getCellValue(row, COL, ['곡/수업 맥락 키워드']),
    interpretPoint   : getCellValue(row, COL, ['수업내용 해석 포인트']),
    evalSentence     : getCellValue(row, COL, ['평가서 조합 문장']),
    futureSentence   : getCellValue(row, COL, ['미래지향 문장']),
    keywords         : getCellValue(row, COL, ['자동화 검색 키워드']),
    strongCandidates : getCellValue(row, COL, ['강매칭 후보']),
    subCandidates    : getCellValue(row, COL, ['보조매칭 후보']),
    memo             : getCellValue(row, COL, ['메모']),
  })).filter(r => r.step !== '');
}

function getCellValue(row, COL, names) {
  for (const name of names) {
    if (COL[name] !== undefined) return String(row[COL[name]] || '').trim();
  }
  return '';
}

// ───── 커리큘럼 용어 해석 매칭 ─────
function matchCurriculum(lessons, curriculum) {
  const allLessonText = lessons.map(l => [
    l.content,
    l.song,
    l.homework,
    l.eval,
  ].join(' ')).join(' ');
  const normalizedLessonText = normalizeMatchText(allLessonText);

  // 목적: Step 추정이 아니라, 출석부에 적힌 용어를 커리큘럼 쉬운해석으로 풀어내기 위한 자료를 찾는다.
  const scores = curriculum.map(step => {
    const strongTerms = splitTerms([
      step.keywords,
      step.strongCandidates,
      step.relatedDrill,
      step.relatedPattern,
    ].join(','));
    const subTerms = splitTerms([
      step.subCandidates,
      step.contextKeywords,
      step.basicArea,
      step.sequenceArea,
      step.stepName,
    ].join(','));

    const strongMatches = findMatchedTerms(strongTerms, normalizedLessonText);
    const subMatches = findMatchedTerms(subTerms, normalizedLessonText);
    const score = strongMatches.length * 3 + subMatches.length;

    return {
      step,
      score,
      strongMatches,
      subMatches,
      allMatches: [...strongMatches, ...subMatches],
    };
  }).filter(s => s.score > 0);

  scores.sort((a, b) => b.score - a.score);

  return scores.slice(0, 4).map(s => ({
    ...s.step,
    score: s.score,
    matchedKeywords: uniqueList(s.allMatches).slice(0, 10),
    strongMatches: uniqueList(s.strongMatches).slice(0, 10),
    subMatches: uniqueList(s.subMatches).slice(0, 10),
  }));
}

function splitTerms(text) {
  return String(text || '')
    .split(/[,，\n]+/)
    .map(t => t.trim())
    .filter(Boolean);
}

function findMatchedTerms(terms, normalizedLessonText) {
  return terms.filter(term => {
    const normalizedTerm = normalizeMatchText(term);
    return normalizedTerm && normalizedLessonText.includes(normalizedTerm);
  });
}

function uniqueList(items) {
  return [...new Set(items.filter(Boolean))];
}

// ───── 매칭용 텍스트 정규화 ─────
function normalizeMatchText(text) {
  return String(text || '')
    .toLowerCase()
    .replace(/플로우/g, 'cell')
    .replace(/셀/g, 'cell')
    .replace(/드릴/g, 'drill')
    .replace(/그루브/g, 'groove')
    .replace(/필인/g, 'fill')
    .replace(/번/g, '')
    .replace(/#/g, '')
    .replace(/\s+/g, '')
    .replace(/[()\[\]{}'"“”‘’.,，:：;；/\\|_\-–—]/g, '')
    .trim();
}

// ───── Claude API 호출 ─────
function callOpenAIAPI(lessons, matchedSteps, curriculum) {
  const apiKey = getApiKey();

  const lessonSummary = lessons.map((l, i) =>
    `[${i + 1}회] 날짜:${l.date} 수업내용:${l.content} 수업곡:${l.song} 과제:${l.homework} 기존평가:${l.eval}`
  ).join('\n');

  const termExplanationSummary = matchedSteps.map(s =>
    `매칭단계: ${s.step} ${s.stepName}\n출석부에서 발견된 용어: ${(s.matchedKeywords || []).join(', ') || '없음'}\n기본기 용어 해석: ${s.basicEasy}\n시퀀스/셀/그루브/필인 해석: ${s.sequenceEasy}\n수업내용 해석 포인트: ${s.interpretPoint}\n평가서에 풀어쓸 문장: ${s.evalSentence}\n다음 수업 방향 문장: ${s.futureSentence}`
  ).join('\n\n');

  const prompt = `당신은 드럼 레슨 기록을 학부모와 성인회원 모두에게 자연스러운 짧은 레슨 리포트로 바꾸는 작성자입니다.
핵심은 많이 설명하는 것이 아니라, 수업기록에 적힌 각 연습을 정확히 구분하고 "무엇을 하고 있는지, 무엇을 위해 하고 있는지, 다음에 무엇을 이어갈지"를 함축적으로 정리하는 것입니다.

【최근 수업기록】
${lessonSummary}

【커리큘럼 용어 해석 자료】
${termExplanationSummary || '(매칭된 용어 해석 자료 없음)'}

【출력 형식】
- 총 2문단만 작성합니다.
- 각 문단은 1문장만 작성합니다.
- 전체 150~230자 내외로 작성합니다.
- 첫 문단: 수업에서 다룬 개별 연습들과 각 연습의 목적.
- 둘째 문단: 실제 변화 또는 현재 병목 1개 + 다음 수업 방향 1개.
- 평가 초안 본문만 출력합니다.

【정확도 규칙】
- 수업기록의 연습들이 서로 다른 개별 연습이면 억지로 하나의 흐름으로 묶지 마세요.
- 서로 관련 없는 연습을 "함께", "통해", "적용하여", "중심으로" 같은 말로 연결하지 마세요.
- 개별 연습은 "A를 하며 ~을 준비하고, B에서는 ~을 연습하고 있습니다"처럼 병렬로 설명하세요.
- 곡 연습은 테크닉 연습의 결과물로 단정하지 말고, 곡 안에서 따로 다룬 적용 연습으로 설명하세요.
- 다음 수업 방향은 모든 항목을 다 묶지 말고, 기록상 가장 중요한 병목 1개만 선택하세요.

【존댓말 규칙】
- 반드시 정중한 존댓말로 작성합니다.
- 모든 문장은 "습니다", "합니다", "필요합니다", "예정입니다", "지도하겠습니다" 중 하나처럼 존댓말로 끝냅니다.
- "했다", "집중했다", "필요하다", "맞춰야 한다", "초점을 맞춰야 한다" 같은 반말/보고서체 종결은 절대 쓰지 마세요.

【작성 방식】
- 학생 이름, 회원 이름을 쓰지 마세요.
- "학생은", "회원은", "수강생은" 같은 주어를 쓰지 마세요.
- "최근 수업에서는", "최근 4회의 수업에서", "이번 기간 동안", "수업을 통해"로 시작하지 마세요.
- "기본 드럼 패턴", "기본 리듬", "기본 연습"처럼 넓고 디테일 없는 표현을 쓰지 마세요.
- 용어를 사전처럼 정의하지 말고, 실제 수업에서 하고 있는 연습과 목적이 자연스럽게 드러나게 쓰세요.
- 대표 곡명이나 대표 용어는 꼭 필요할 때 1개만 남기세요.

【금지】
- 3문단 이상
- 5문장 이상
- "리듬감이 좋아졌습니다", "안정적인 연주력이 기대됩니다" 같은 일반 표현
- "기본 드럼 패턴", "기본 리듬", "기본기"만 단독으로 쓰는 표현
- 서로 관련 없는 연습을 하나의 흐름처럼 엮는 표현
- 수업기록에 없는 칭찬
- 제목, 번호, 분석 과정`;

  const payload = {
    model: 'claude-sonnet-4-6',
    max_tokens: 320,
    temperature: 0.1,
    messages: [
      {
        role: 'user',
        content: prompt,
      },
    ],
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
  const code = response.getResponseCode();
  const rawText = response.getContentText();
  const body = JSON.parse(rawText);

  if (code < 200 || code >= 300) {
    throw new Error(`API 오류 (${code}): ${body.error?.message || rawText}`);
  }

  const outputText = extractClaudeText(body);
  if (!outputText) {
    throw new Error('API 응답에서 평가 초안 텍스트를 찾을 수 없습니다.');
  }

  return outputText.trim();
}

// ───── Claude 응답 텍스트 추출 ─────
function extractClaudeText(body) {
  if (!body.content || !Array.isArray(body.content)) return '';

  const parts = [];
  body.content.forEach(content => {
    if (content.type === 'text' && content.text) {
      parts.push(String(content.text));
    }
  });

  return parts.join('\n').trim();
}
