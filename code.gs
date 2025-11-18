// =================================================================
// 1. 최상위 설정 및 전역 변수
// =================================================================
const SHEET_ID = ''; // Q&A와 챗봇이 함께 사용할 시트 ID
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

const DEV_TEST_MODEL = 'gemini-2.5-flash'; // 개발 및 테스트용 모델
const PRODUCTION_MODEL = 'gemini-2.5-flash'; // 실제 운영용 모델
const CURRENT_GEMINI_MODEL = DEV_TEST_MODEL; // <-- 실제 운영 시 PRODUCTION_MODEL로 변경

const SHEET_NAME = '상담로그';
const RAG_SHEET_NAME = 'RAG_DB';
const QNA_SHEET_NAME = 'QNA'; // Q&A 게시판이 사용할 시트 이름
const DEBUG_MODE = true; // 디버깅이 필요할 때 true로 변경
let debugLogs = [];

const DEPARTMENT_LIST = [
  '간호학과', '스마트팜식품융합과', '반려동물과', '사회복지과',
  '소방안전관리과', '외식창업조리과', '유아교육과', '유통경영과',
  '임상병리과', '작업치료과', '전기과', '제과제빵과', '치기공과', '치위생과',
  '한국어과', '호텔관광서비스과', '호텔조리계열', '언어치료과', '보건의료행정과',
  '치위생학과(전공심화)', '유아교육학과(전공심화)', '제과제빵학과(전공심화)',
  '한식조리과(조기취업형)', '중식조리과(조기취업형)', '일식조리과(조기취업형)', '서양식조리과(조기취업형)', '베이커리카페과(조기취업형)'
];

// ... (DEPARTMENT_PHONE_NUMBERS, BOT_GREETING 등 나머지 전역 변수들은 기존과 동일하게 유지)


// =================================================================
// 2. 메인 진입점 함수 (웹앱 라우팅 및 페이지 로드)
// =================================================================
function doGet(e) {
  // 관리자 페이지 라우팅
  if (e.parameter.page === 'admin') {
    return HtmlService.createTemplateFromFile('admin.html').evaluate()
      .setTitle('관리자 페이지')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  // Q&A 게시판 라우팅
  if (e.parameter.page === 'qna') {
    return HtmlService.createTemplateFromFile('qna_board.html').evaluate()
      .setTitle('Q&A 게시판')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  // 기본 챗봇 페이지
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('AI 학사상담')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}


/* ================================================================== */
/* 3. 공개 Q&A 게시판 기능                                            */
/* ================================================================== */
function addQuestion(formData) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(QNA_SHEET_NAME);
    const userKey = Session.getTemporaryActiveUserKey() || 'anonymous';
    sheet.appendRow([
      new Date(),
      formData.name,
      formData.title,
      formData.question,
      "", // Answer (초기값)
      userKey // SessionID
    ]);
    return { success: true };
  } catch (e) {
    console.error("addQuestion 오류: " + e.toString());
    return { success: false, message: "문의 등록 중 오류가 발생했습니다: " + e.message };
  }
}

// =================================================================
// 4. 챗봇 기능 (기존 코드)
// =================================================================

/**
 * 챗봇 UI에 학과 목록을 제공합니다.
 * @returns {Array<string>} 학과 이름 배열
 */
function getDepartmentList() {
  return DEPARTMENT_LIST;
}

/**
 * 챗봇 질문 처리의 메인 로직입니다.
 * @param {Object} data 사용자 입력 데이터
 * @returns {Object} AI 답변 또는 오류 객체
 */
function processQuestion(data) {
  try {
    const startTimeObject = new Date(); 
    const formattedTimestamp = Utilities.formatDate(startTimeObject, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

    debugLogs = []; 
    debugLog(`--- 질문 처리 시작 (Timestamp: ${formattedTimestamp}) ---`);

    const validation = validateInput(data);
    if (!validation.isValid) {
      return { success: false, error: validation.message, debugLogs: getDebugLogs() };
    }

    debugLog(`원본 질문: "${data.question}", 카테고리: "${data.category || '없음'}", 세션ID: ${data.sessionId}`);

    // 1. 사용자 지정 언어
    const detectedLang = data.selectedLang || 'ko';
    debugLog(`사용자 선택 언어: ${detectedLang}`);

    // 2. 질문을 한국어로 변환 (사용자 선택 언어 → 한국어)
    let translatedQuestion = data.question;
    if (detectedLang !== 'ko') {
      translatedQuestion = LanguageApp.translate(data.question, detectedLang, 'ko');
      debugLog(`질문 한국어 번역: "${translatedQuestion}"`);
    }

    // 3. RAG 검색
    const ragResults = searchRAGData(translatedQuestion, data.category, data.admissionYear);
    debugLog(`RAG DB 검색 결과: ${ragResults.length}개 문서 발견.`);

    // 4. AI 응답 (한국어)
    const aiResponseInKorean = getAiResponse({
      admissionYear: data.admissionYear,
      department: data.department,
      category: data.category,
      question: translatedQuestion
    }, ragResults);
    debugLog(`AI 한국어 답변: "${aiResponseInKorean}"`);

    // 5. 사용자 선택 언어로 번역
    let finalAiResponse = aiResponseInKorean;
    if (detectedLang !== 'ko') {
      finalAiResponse = LanguageApp.translate(aiResponseInKorean, 'ko', detectedLang);
      debugLog(`AI 답변 ${detectedLang}로 번역: "${finalAiResponse}"`);
    }

    const endTime = new Date();
    const responseTime = Math.round((endTime.getTime() - startTimeObject.getTime()) / 1000);

    // 6. 로그 저장
    saveLogToSheet({
      timestamp: formattedTimestamp,
      admissionYear: data.admissionYear,
      department: data.department,
      category: data.category,
      originalQuestion: data.question,
      translatedQuestion: translatedQuestion,
      aiResponseInKorean: aiResponseInKorean,
      finalAnswer: finalAiResponse,
      responseTime: responseTime,
      ragUsed: (ragResults && ragResults.length > 0) ? 'Y' : 'N',
      sessionId: data.sessionId,
      detectedLanguage: detectedLang
    });

    debugLog('--- 질문 처리 완료 ---');

    return {
      success: true, answer: finalAiResponse, responseTime: responseTime,
      debugLogs: getDebugLogs()
    };
  } catch (error) {
    console.error('Error in processQuestion:', error.stack);
    const errorMessage = error.message || '처리 중 알 수 없는 오류가 발생했습니다.';
    const userFriendlyError = '죄송합니다. AI 서비스에 일시적인 문제가 있습니다. 잠시 후 다시 시도해 주세요.';

    debugLog(`최종 오류 발생: ${errorMessage}`);
    debugLog('--- 질문 처리 실패 ---');

    return { success: false, error: userFriendlyError, debugLogs: getDebugLogs() };
  }
}


// =================================================================
// 5. AI 응답 생성 및 API 호출 (기존 코드)
// =================================================================

function getAiResponse(data, ragResults) {
  if (!GEMINI_API_KEY) {
    debugLog('오류: 스크립트 속성에 Gemini API 키가 설정되지 않았습니다.');
    throw new Error('서버 관리자에게 문의하세요: API 키가 설정되지 않았습니다.');
  }

  const prompt = createEnhancedPrompt(data, ragResults);
  debugLog("--- 최종 생성된 프롬프트 ---");
  debugLog(prompt);
  debugLog("--------------------------");

  return callGeminiWithRetry(prompt);
}

function callGeminiWithRetry(prompt) {
  const MAX_RETRIES = 3;
  let waitTime = 1000;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CURRENT_GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.4, topK: 40, topP: 0.95, maxOutputTokens: 1024 }
  };
  const options = {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  for (let i = 0; i < MAX_RETRIES; i++) {
    debugLog(`Gemini API 호출 시도 (${i + 1}/${MAX_RETRIES})...`);
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      try {
        const responseData = JSON.parse(responseBody);
        
        // 👇👇👇 이 부분이 수정되었습니다. 👇👇👇
        if (responseData.candidates && responseData.candidates.length > 0 && 
            responseData.candidates[0].content && responseData.candidates[0].content.parts &&
            responseData.candidates[0].content.parts.length > 0) {
          debugLog('Gemini API 호출 성공!');
          return responseData.candidates[0].content.parts[0].text;
        } else {
          // 응답은 200이지만 candidates가 없거나 비어있는 경우 (예: 안전 필터로 인해 거부된 경우)
          const errorDetail = responseData.promptFeedback ? `(사유: ${responseData.promptFeedback.blockReason})` : '';
          debugLog(`API 응답 형식 오류: ${responseBody}`);
          throw new Error(`API 응답 데이터 구조에 문제가 있습니다 ${errorDetail}. 응답 본문을 확인하세요.`);
        }
        // 👆👆👆 수정된 부분 끝 👆👆👆
        
      } catch (e) {
        debugLog(`API 응답 JSON 파싱 실패: ${e.message}`);
        throw new Error('API 응답을 처리하는 데 실패했습니다.');
      }
    }

    if (responseCode === 503 || responseCode === 500 || responseCode === 429) {
      debugLog(`API 오류 (코드: ${responseCode}), ${waitTime / 1000}초 후 재시도합니다.`);
      if (i < MAX_RETRIES - 1) {
        Utilities.sleep(waitTime);
        waitTime *= 2;
      }
    } else {
      debugLog(`복구 불가능한 API 오류 발생 (코드: ${responseCode}): ${responseBody}`);
      throw new Error(`API 요청 실패 (코드: ${responseCode})`);
    }
  }

  debugLog('최대 재시도 횟수를 초과하여 API 호출에 최종적으로 실패했습니다.');
  throw new Error('AI 서비스가 응답하지 않습니다. 잠시 후 다시 시도해주세요.');
}

function createEnhancedPrompt(data, ragResults) {
  const { 
    admissionYear = '정보 없음', 
    department = '정보 없음', 
    category = '정보 없음', 
    question = '' 
  } = data;

  let ragContext = '';
  const foundCategories = new Set(); 
  if (ragResults && ragResults.length > 0) {
    ragContext = '\n\n=== 참고할 학사 규정 및 정보 ===\n';
    ragResults.forEach((result, index) => {
      ragContext += `\n[정보 ${index + 1} | 출처 카테고리: ${result.category} | 제목: ${result.originalTitle}]\n${result.originalContent}\n`;
      if (result.category) {
        foundCategories.add(result.category);
      }
    });
    ragContext += '\n=== 정보 끝 ===\n';
  }

  const nowInKorea = new Date();
  const currentDate = Utilities.formatDate(nowInKorea, "Asia/Seoul", "yyyy년 M월 d일");
  const currentYear = Utilities.formatDate(nowInKorea, "Asia/Seoul", "yyyy");
  
  let uncertaintyGuideline = '5. **불확실성**: 정보가 없거나 확실하지 않으면 추측하지 말고, "다른 단어들을 사용해서 다시 질문해 보세요."라고 안내하세요.';

const answerGuideline = `## 답변 가이드라인
1.  **인사말 규칙**: "안녕하세요!"와 같이 정중하고 일반적인 인사말로 시작하세요. 절대로 학생 이름을 만들지 마세요.
2.  **시점 판단**: 현재는 '${currentDate}' 입니다. 질문에 연도가 없으면, **${currentYear}년** 기준으로 답변하세요.
3.  **학기 판단 규칙**: '이번 학기'와 같은 표현이 있으면, 현재 날짜를 기준으로 다음 규칙에 따라 학기를 먼저 판단하세요.
    -   **1학기**: 매년 3월 1일 ~ 8월 31일
    -   **2학기**: 매년 9월 1일 ~ 다음 해 2월 말일
    -   **(예시)** 현재가 2025년 10월 17일이므로, '이번 학기 방학'은 **2025학년도 2학기 동계 방학**을 의미합니다. 1학기 정보는 참고용으로만 제시하세요.
4.  **최우선 참고**: "참고할 학사 규정 및 정보"가 있다면, 그 내용을 기반으로 답변하세요.
5.  **구조화된 형식**: 답변을 명확하게 전달하기 위해 제목과 글머리 기호를 사용하여 구조화하세요. 주요 섹션에는 제목(예: ### 제목)을 사용하고, 각 항목은 글머리 기호(-)로 나열하세요. 필요하다면 하위 항목을 들여쓰기하여 추가 설명을 제공할 수 있습니다.
    (예시)
    ### 증명서 발급 방법
    - **온라인 발급**
        - 혜전대학교 홈페이지 통합정보시스템에 로그인 후, '증명서 발급' 메뉴에서 신청 가능합니다.
    - **무인 발급기 이용**
        - 교내 학생회관 1층에 설치된 무인 발급기를 통해 즉시 발급받을 수 있습니다.
${uncertaintyGuideline}
7.  **어조 및 형식**: 전문적이고 친절한 톤을 유지하고, **반드시 한국어로** 가독성 좋게 작성하세요. (최종 답변 번역은 시스템이 처리합니다.)`;

  return `당신은 혜전대학교의 전문적이고 친절한 학사상담 AI 챗봇입니다.

## 시스템 현재 시점
- 현재 날짜: ${currentDate}

## 학생 정보
- 입학연도: ${admissionYear}년
- 학과: ${department}
- 질문 분야: ${category}

## 학생 질문
"${question}"
${ragContext}
${answerGuideline}

위 가이드라인에 따라 학생의 질문에 답변해주세요.
답변:`;
}

// =================================================================
// 6. RAG DB 검색 및 유사도 계산 (기존 코드)
// =================================================================

function searchRAGData(originalQuery, category, admissionYear) {
  try {
    const allRagData = loadRAGToMemory();
    if (!allRagData || allRagData.length === 0) return [];

    const queryKeywords = extractKeywords(originalQuery);
    debugLog(`질문에서 추출된 핵심 키워드: [${queryKeywords.join(', ')}]`);

    if (queryKeywords.length === 0) {
      debugLog('질문에서 유효한 키워드를 찾지 못해 검색을 종료합니다.');
      return [];
    }

    const scoredDocuments = allRagData.map(doc => {
      const { score, log } = calculateRelevanceScore(queryKeywords, doc, admissionYear, category);
      if (score > 0) {
        debugLog(log);
      }
      return { ...doc, similarity: score };
    });

    const MINIMUM_SCORE = 10; // 관련 문서로 판단할 최소 점수
    return scoredDocuments
      .filter(doc => doc.similarity >= MINIMUM_SCORE)
      .sort((a, b) => b.similarity - a.similarity)
      .slice(0, 5); // 상위 5개 결과만 반환

  } catch (error) {
    console.error('RAG 검색 오류:', error.stack);
    debugLog(`RAG 검색 중 오류 발생: ${error.message}`);
    return [];
  }
}

function extractKeywords(text) {
  const keywordDictionary = [
    '학사일정', '개강', '방학', '기말고사', '수강신청', '휴학', '복학', '졸업', '학점', '등록', '등록금',
    '현장실습', '인턴', '학제', '입학정원', '전화번호', '연락처', '개교기념일', '공휴일', '장학', '장학금',
    '성적', '평가', '비율', '등급', '평점', '기준', '성적평가비율', '이사장장학금', '총장장학금', 
    '수석장학금', '우수장학금','영어능력우수', '토익점핑', '공로장학금', '목련장학금', '혜전동문', 
    '교직원장학금', '다문화장학금', '성인학습자', '혜전드림', '혜전생활', '재난지원', '학과장추천', 
    '지역인재', '향토지역인재', '혜전홍성', '면접위주', '교육협력고', '마일리지', '해외연수', 
    '총장특별', '전공심화', '산업체위탁', '보훈장학금', '통일부장학금', '단곡장학금', '명예총장', '국가근로',
    '개강일', '여름방학', '겨울방학', '중간고사', '기말', '중간', '경고', '전과', '복학', '성적평가', '성적등급', 
    '도서관', '도서관 이용', 'DVD', '도서대출', '도서관대출', '졸업학점',
    '강의평가', '계약학과', '계절학기', '명예졸업', '사회봉사', '산업체 위탁교육', '소수집단', '다문화',  '장애학생', '외국인 유학생',
    '수업운영', '외국대학 연수', '교환학생', '유학생 입학' , '유학생 신입학', '유학생 편입학', '원격수업', '장애학생 지원',
    '전공선택',  '조기취업자', '조기취업', '졸업시험', '집중이수제', '전공심화', '학생생활', '생활관', '기숙사',
    '학위', '학위종류', '학점 인정'
  ];

  const foundKeywords = new Set();
  const normalizedText = text.replace(/\s+/g, '').toLowerCase(); // 공백 제거 및 소문자화

  keywordDictionary.forEach(kw => {
    if (normalizedText.includes(kw.toLowerCase())) {
      foundKeywords.add(kw);
    }
  });

  return [...foundKeywords];
}

function calculateRelevanceScore(queryKeywords, doc, admissionYear, selectedCategory) {
  let score = 0;
  const matchedLog = [];

  const docTitle = doc.searchTitle || '';
  const docContent = doc.searchContent || '';

  // 1. 키워드 매칭 점수
  queryKeywords.forEach(keyword => {
    if (docTitle.includes(keyword)) {
      score += 30;
      matchedLog.push(`'${keyword}'(제목)`);
    } 
    else if (docContent.includes(keyword)) {
      score += 10;
      matchedLog.push(`'${keyword}'(본문)`);
    }
  });

  // 2. 특별 조건 보너스 점수
  if (selectedCategory && doc.category && selectedCategory === doc.category) {
    score += 50;
    matchedLog.push('[카테고리 일치 +50]');
  }

  if (selectedCategory === '졸업학점, 성적, 학점' && admissionYear && docContent.includes(admissionYear)) {
    score += 200;
    matchedLog.push(`[★학번(${admissionYear}) 일치 +200]`);
  }
  
  const logMessage = `   - [${doc.originalTitle}] 유사도: [${matchedLog.join(', ')}] >> 최종 점수: ${score}`;
  
  return { score, log: logMessage };
}

// =================================================================
// 7. 유틸리티 및 헬퍼 함수 (로깅, 캐시, 유효성 검사 등)
// =================================================================

function debugLog(message) {
  if (DEBUG_MODE) {
    const timestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "HH:mm:ss");
    const logMessage = `[DEBUG ${timestamp}] ${message}`;
    console.log(logMessage);
    debugLogs.push(logMessage);
  }
}

function getDebugLogs() {
  if (DEBUG_MODE) {
    return [...debugLogs];
  }
  return [];
}

function loadRAGToMemory() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get('ragData');

  if (cachedData) {
    debugLog('CacheService에서 RAG 데이터를 로드했습니다.');
    return JSON.parse(cachedData);
  }

  debugLog('CacheService에 데이터가 없어, Google Sheets에서 RAG DB를 새로 로드합니다...');
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(RAG_SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
    
    const ragData = data.map(row => ({
      id: row[0], originalTitle: row[1], searchTitle: row[2], 
      originalContent: row[3], searchContent: row[4], category: row[5], createdAt: row[6]
    }));
    
    cache.put('ragData', JSON.stringify(ragData), 1800); // 30분간 캐시
    debugLog(`${ragData.length}개의 문서를 메모리에 로드하고, CacheService에 30분간 저장했습니다.`);
    
    return ragData;
  } catch (error) {
    console.error('RAG 메모리/캐시 로드 오류:', error.stack);
    debugLog(`RAG 로드 중 오류 발생: ${error.message}`);
    return [];
  }
}

function saveLogToSheet(logData) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, 12).setValues([
        ['타임스탬프', '입학연도', '학과', '주제분류', '원본질문내용', '번역된질문내용', 'AI한국어답변', '최종답변', '응답시간(초)', 'RAG활용여부', '세션ID', '감지된언어']
      ]);
    }
    sheet.appendRow([
      logData.timestamp, logData.admissionYear, logData.department, logData.category,
      logData.originalQuestion, logData.translatedQuestion, logData.aiResponseInKorean, logData.finalAnswer,
      logData.responseTime, logData.ragUsed, logData.sessionId, logData.detectedLanguage
    ]);
  } catch (error) {
    console.error('Error saving log to sheet:', error.stack);
  }
}

function validateInput(data) {
  if (!data.admissionYear || !data.department || !data.category || !data.question) {
    return { isValid: false, message: '모든 필드를 입력해 주세요.' };
  }
  if (!/^\d{4}$/.test(data.admissionYear) || +data.admissionYear < 2000 || +data.admissionYear > new Date().getFullYear() + 1) {
    return { isValid: false, message: '올바른 입학연도를 입력해 주세요.' };
  }
  if (data.question.trim().length < 5) {
    return { isValid: false, message: '질문은 5자 이상으로 입력해 주세요.' };
  }
  return { isValid: true };
}

function checkDuplicateQuestion(currentQuestion, sessionId, startTime) {
  // 이 함수는 현재 로직에서 직접 호출되지 않으므로, 필요 시 활성화하여 사용 가능합니다.
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    if (sheet.getLastRow() <= 1) return false;

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();
    const fiveMinutesAgo = new Date(startTime.getTime() - 5 * 60 * 1000);

    const currentKeywords = extractKeywords(currentQuestion);
    if (currentKeywords.length === 0) return false;

    for (let i = data.length - 1; i >= 0; i--) {
      const rowTime = new Date(data[i][0]);
      if (rowTime < fiveMinutesAgo) break;

      const logSessionId = data[i][10]; 
      if (logSessionId === sessionId) {
        const logQuestion = data[i][4]; 
        const logKeywords = extractKeywords(logQuestion);
        
        const intersection = currentKeywords.filter(kw => logKeywords.includes(kw));
        const similarity = intersection.length / Math.max(currentKeywords.length, logKeywords.length);

        if (similarity >= 0.9) {
          debugLog(`중복 질문 감지: 동일 세션(${sessionId})에서 키워드 유사도 ${similarity.toFixed(2)}의 질문 발견.`);
          return true;
        }
      }
    }
    return false;
  } catch (error) {
    console.error('Error checking duplicate question:', error.stack);
    return false;
  }
}


// =================================================================
// 8. [관리자용] 캐시 관리 함수 (기존 코드)
// =================================================================

/**
 * 관리자가 RAG DB 변경 후 수동으로 캐시를 삭제할 때 사용합니다.
 */
function clearRAGCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove('ragData');
    
    SpreadsheetApp.getUi().alert('성공!', 'RAG 데이터 캐시를 성공적으로 삭제했습니다. 다음 질문부터는 수정된 DB가 즉시 반영됩니다.', SpreadsheetApp.getUi().ButtonSet.OK);
    console.log('RAG 데이터 캐시가 성공적으로 삭제되었습니다.');
    
  } catch (error) {
    console.error('캐시 삭제 중 오류 발생:', error.stack);
    SpreadsheetApp.getUi().alert('오류', `캐시 삭제 중 오류가 발생했습니다: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


/* ================================================================== */
/* 9. 관리자 페이지 기능 (QNA, 통계, RAG 통합)                       */
/* ================================================================== */

/**
 * 관리자 로그인 자격 증명 확인
 */
function checkAdminCredentials(credentials) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('관리자');
    if (!sheet) {
      if (credentials.id === 'admin' && credentials.password === 'admin123') return { success: true };
      return { success: false, message: "관리자 설정이 필요합니다." };
    }
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] == credentials.id && data[i][1] == credentials.password) return { success: true };
    }
    return { success: false, message: "아이디 또는 비밀번호가 일치하지 않습니다." };
  } catch (e) {
    console.error("checkAdminCredentials 오류: " + e.toString());
    return { success: false, message: "로그인 중 오류가 발생했습니다." };
  }
}

/**
 * 관리자용 Q&A 데이터 조회
 */
function getQnaDataForAdmin() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(QNA_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5);
    const data = dataRange.getValues();
    return data.map((row, index) => {
      const timestamp = row[0];
      const isDate = timestamp instanceof Date && !isNaN(timestamp);
      return {
        rowIndex: index + 2,
        timestamp: isDate ? timestamp.toISOString() : new Date().toISOString(),
        name: row[1], title: row[2], question: row[3], answer: row[4]
      };
    }).reverse();
  } catch (e) {
    console.error("getQnaDataForAdmin 오류: " + e.toString());
    return [];
  }
}

/**
 * Q&A 답변 업데이트
 */
function updateAnswer(data) {
  try {
    const { rowIndex, answer } = data;
    if (!rowIndex || typeof answer === 'undefined') return { success: false, message: "잘못된 요청입니다." };
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(QNA_SHEET_NAME);
    sheet.getRange(rowIndex, 5).setValue(answer);
    return { success: true };
  } catch (e) {
    console.error("updateAnswer 오류: " + e.toString());
    return { success: false, message: "답변 저장 중 오류가 발생했습니다." };
  }
}

/**
 * 통계 데이터 집계 (수정됨)
 */
function getStatisticsData() {
  try {
    const 상담로그_시트 = SpreadsheetApp.openById(SHEET_ID).getSheetByName('상담로그');
    const QNA_시트 = SpreadsheetApp.openById(SHEET_ID).getSheetByName('QNA');

    const accessCounts = { today: 0, week: 0, month: 0, year: 0 };
    const categoryCounts = {};
    const departmentCounts = {};
    
    // 1. 챗봇 이용 현황 집계 (상담로그 시트)
    if (상담로그_시트 && 상담로그_시트.getLastRow() > 1) {
        const now = new Date();
        const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
        const weekStart = new Date(now.setDate(now.getDate() - now.getDay()));
        weekStart.setHours(0, 0, 0, 0);
        const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
        const yearStart = new Date(now.getFullYear(), 0, 1);

        const logData = 상담로그_시트.getRange(2, 1, 상담로그_시트.getLastRow() - 1, 4).getValues();

        logData.forEach(row => {
            const timestamp = new Date(row[0]);
            const department = row[2];
            const category = row[3];

            if (timestamp >= todayStart) accessCounts.today++;
            if (timestamp >= weekStart) accessCounts.week++;
            if (timestamp >= monthStart) accessCounts.month++;
            if (timestamp >= yearStart) accessCounts.year++;

            if (category) categoryCounts[category] = (categoryCounts[category] || 0) + 1;
            if (department) departmentCounts[department] = (departmentCounts[department] || 0) + 1;
        });
    }

    // 2. Q&A 처리 현황 집계 (QNA 시트)
    let qnaStats = { total: 0, answered: 0, unanswered: 0 };
    if (QNA_시트 && QNA_시트.getLastRow() > 1) {
        const qnaData = QNA_시트.getRange(2, 5, QNA_시트.getLastRow() - 1, 1).getValues();
        qnaStats.total = qnaData.length;
        qnaStats.answered = qnaData.filter(row => row[0] && String(row[0]).trim() !== '').length;
        qnaStats.unanswered = qnaStats.total - qnaStats.answered;
    }

    return {
        accessCounts,
        qnaStats,
        categoryCounts,
        departmentCounts
    };

  } catch (e) {
    console.error("getStatisticsData 오류: " + e.toString());
    // 오류 발생 시 클라이언트에서 에러를 처리할 수 있도록 null을 반환합니다.
    return null;
  }
}

/**
 * RAG 카테고리 목록 조회
 */
function getRAGCategories() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(RAG_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 6, sheet.getLastRow() - 1, 1).getValues();
    const categories = new Set(data.map(row => row[0]).filter(Boolean));
    return Array.from(categories).sort();
  } catch (e) {
    console.error("getRAGCategories 오류: " + e.toString());
    return [];
  }
}

/**
 * RAG 전체 데이터 삭제
 */
function deleteAllRAGData() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(RAG_SHEET_NAME);
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
    return { success: true };
  } catch (e) {
    console.error("deleteAllRAGData 오류: " + e.toString());
    return { success: false, message: e.message };
  }
}

/**
 * RAG 특정 카테고리 데이터 삭제
 */
function deleteRAGCategory(categoryToDelete) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(RAG_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const rowsToDelete = [];
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][5] === categoryToDelete) rowsToDelete.push(i + 1);
    }
    if (rowsToDelete.length > 0) rowsToDelete.forEach(rowIndex => sheet.deleteRow(rowIndex));
    return { success: true };
  } catch (e) {
    console.error("deleteRAGCategory 오류: " + e.toString());
    return { success: false, message: e.message };
  }
}

/**
 * RAG 새 카테고리 추가
 */
function addEmptyRAGEntryWithCategory(categoryName) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(RAG_SHEET_NAME);
    if (getRAGCategories().includes(categoryName)) {
      return { success: false, message: "이미 존재하는 카테고리입니다." };
    }
    const newId = "RAG-" + new Date().getTime();
    sheet.appendRow([ newId, `${categoryName} 관련 규정 (제목)`, "", `${categoryName}에 대한 내용을 여기에 입력하세요.`, "", categoryName, new Date() ]);
    return { success: true };
  } catch (e) {
    console.error("addEmptyRAGEntryWithCategory 오류: " + e.toString());
    return { success: false, message: e.message };
  }
}


