const STORAGE_KEYS = {
  settings: "englishGameSettings",
  progress: "englishGameProgress",
};

const SUPPORTED_GRADES = [1, 2, 3, 4, 5, 6];
const GRADE_LABELS = {
  1: "中1",
  2: "中2",
  3: "中3",
  4: "高1",
  5: "高2",
  6: "高3",
};

const POS_LABEL = {
  NOUN: "名詞",
  PROPN: "固有名詞",
  PRON: "代名詞",
  REL_PRON: "関係代名詞",
  INT_PRON: "疑問代名詞",
  VERB: "動詞",
  VERB_ING: "動名詞",
  AUX: "助動詞",
  ADJ: "形容詞",
  ADV: "副詞",
  DET: "限定詞",
  NUM: "数詞",
  ADP: "前置詞",
  PART: "不変化詞",
  PARTICLE: "小詞",
  TO: "to不定詞",
  CCONJ: "等位接続詞",
  SCONJ: "従位接続詞",
  REL_ADV: "関係副詞",
  INT_ADV: "疑問副詞",
  INTJ: "間投詞",
  PUNCT: "句読点",
};

const POS_COLOR_CLASS = {
  NOUN: "noun",
  PROPN: "noun",
  PRON: "pronoun",
  REL_PRON: "pronoun",
  INT_PRON: "pronoun",
  VERB: "verb",
  VERB_ING: "verb",
  AUX: "auxiliary",
  ADJ: "adjective",
  ADV: "adverb",
  DET: "article",
  NUM: "adjective",
  ADP: "preposition",
  PART: "adverb",
  PARTICLE: "verb",
  TO: "preposition",
  CCONJ: "conjunction",
  SCONJ: "conjunction",
  REL_ADV: "adverb",
  INT_ADV: "adverb",
  INTJ: "adverb",
};

const STANZA_LABEL_JA = {
  root: "",
  nsubj: "",
  "nsubj:pass": "",
  csubj: "",
  obj: "",
  iobj: "",
  xcomp: "",
  ccomp: "",
  advcl: "",
  verb_subordinate: "",
  "acl:relcl": "",
  obl: "",
  "obl:agent": "",
  "obl:tmod": "",
  "obl:unmarked": "",
  advmod: "",
  amod: "",
  nummod: "",
  "nmod:poss": "",
  case: "",
  det: "",
  mark: "",
  conj: "",
  cc: "",
  expl: "",
  cop: "",
  appos: "",
  residual: "",
  implicit_subject: "",
  predicate_tail: "",
};

const appState = {
  grade: null,
  questionSet: [],
  currentQuestionIndex: 0,
  correctCount: 0,
  wrongCount: 0,
  wrongQuestionIds: [],
  currentQuestion: null,
  wordBank: [],
  answer: [],
  currentAttemptCount: 0,
  historySnapshots: [],
  wrongPositions: new Set(),
  activeGapIndex: null,
  settings: {
    soundOn: true,
  },
  dataCache: new Map(),
  listView: {
    grade: 1,
    langMode: "en",
  },
};

const ui = {
  headerSubtitle: document.querySelector(".app-header p"),
  screens: {
    home: document.getElementById("home-screen"),
    list: document.getElementById("list-screen"),
    quiz: document.getElementById("quiz-screen"),
    result: document.getElementById("result-screen"),
  },
  gradeGrid: document.getElementById("grade-grid"),
  gradeButtons: [...document.querySelectorAll(".grade-btn")],
  soundToggle: document.getElementById("sound-toggle"),
  questionProgress: document.getElementById("question-progress"),
  scoreMini: document.getElementById("score-mini"),
  currentGradeLabel: document.getElementById("current-grade-label"),
  gradeAchievement: document.getElementById("grade-achievement"),
  homeRateElements: Object.fromEntries(
    SUPPORTED_GRADES.map((grade) => [grade, document.getElementById(`rate-grade-${grade}`)])
  ),
  questionTitle: document.getElementById("question-title"),
  jpMeaning: document.getElementById("jp-meaning"),
  wordBank: document.getElementById("word-bank"),
  answerZone: document.getElementById("answer-zone"),
  feedback: document.getElementById("feedback"),
  structurePanel: document.getElementById("structure-guide"),
  structureSummary: document.getElementById("structure-summary"),
  structureCard: document.getElementById("structure-card"),
  undoBtn: document.getElementById("undo-btn"),
  clearBtn: document.getElementById("clear-btn"),
  checkBtn: document.getElementById("check-btn"),
  nextBtn: document.getElementById("next-btn"),
  quitQuizBtn: document.getElementById("quit-quiz-btn"),
  retryBtn: document.getElementById("retry-btn"),
  backHomeBtn: document.getElementById("back-home-btn"),
  resetProgressBtn: document.getElementById("reset-progress-btn"),
  finalScore: document.getElementById("final-score"),
  finalDetail: document.getElementById("final-detail"),
  reviewList: document.getElementById("review-list"),
  openProblemListBtn: document.getElementById("open-problem-list-btn"),
  backHomeFromListTopBtn: document.getElementById("back-home-from-list-top-btn"),
  backHomeFromListBtn: document.getElementById("back-home-from-list-btn"),
  listGradeButtons: [...document.querySelectorAll(".grade-filter-btn")],
  listLangRadios: [...document.querySelectorAll("input[name='lang-mode']")],
  problemCount: document.getElementById("problem-count"),
  problemList: document.getElementById("problem-list"),
};

let audioCtx = null;
let draggingToken = null;

initialize();

async function initialize() {
  loadSettings();
  await preloadAllGrades();
  applyStaticLabels();
  placeFeedbackNearNextButton();
  bindEvents();
  const singleRequest = getSingleQuestionRequest();
  if (singleRequest) {
    const ok = await startSingleQuestion(singleRequest.grade, singleRequest.id);
    if (ok) {
      return;
    }
  }

  const initialView = getInitialViewRequest();
  if (initialView === "list") {
    showProblemListScreen();
    return;
  }

  showScreen("home");
  refreshAchievementDisplay();
}

function placeFeedbackNearNextButton() {
  if (!ui.feedback || !ui.nextBtn) {
    return;
  }
  const actionRow = ui.nextBtn.parentElement;
  if (!actionRow) {
    return;
  }
  ui.feedback.classList.add("feedback-inline");
  actionRow.insertBefore(ui.feedback, ui.nextBtn);
}

function applyStaticLabels() {
  if (ui.headerSubtitle) {
    ui.headerSubtitle.textContent = "英文並べ替えトレーニング";
  }
  ui.gradeButtons.forEach((button) => {
    const grade = Number(button.dataset.grade);
    const rateNode = button.querySelector("small");
    button.textContent = formatGradeLabel(grade);
    if (rateNode) {
      button.appendChild(rateNode);
    }
  });
  ui.listGradeButtons.forEach((button) => {
    const grade = Number(button.dataset.grade);
    button.textContent = formatGradeLabel(grade);
  });
  if (ui.backHomeFromListTopBtn) {
    ui.backHomeFromListTopBtn.textContent = "\u30c8\u30c3\u30d7\u30da\u30fc\u30b8\u3078\u623b\u308b";
  }
  if (ui.backHomeFromListBtn) {
    ui.backHomeFromListBtn.textContent = "\u30c8\u30c3\u30d7\u30da\u30fc\u30b8\u3078\u623b\u308b";
  }
}

function getSingleQuestionRequest() {
  const params = new URLSearchParams(window.location.search);
  if (params.get("single") !== "1") {
    return null;
  }
  const grade = Number(params.get("grade"));
  const id = String(params.get("id") || "");
  if (!Number.isInteger(grade) || !id) {
    return null;
  }
  return { grade, id };
}


function getInitialViewRequest() {
  const params = new URLSearchParams(window.location.search);
  const view = String(params.get("view") || "").toLowerCase();
  if (view === "list") {
    return "list";
  }
  return "home";
}

function bindEvents() {
  ui.gradeButtons.forEach((button) => {
    button.addEventListener("click", () => startQuiz(Number(button.dataset.grade)));
  });

  ui.soundToggle.addEventListener("change", () => {
    appState.settings.soundOn = ui.soundToggle.checked;
    saveSettings();
  });

  ui.undoBtn.addEventListener("click", undoMove);
  ui.clearBtn.addEventListener("click", clearAnswer);
  ui.checkBtn.addEventListener("click", checkAnswer);
  ui.nextBtn.addEventListener("click", goToNextQuestion);
  ui.quitQuizBtn.addEventListener("click", quitQuizToHome);
  ui.retryBtn.addEventListener("click", () => startQuiz(appState.grade));
  ui.backHomeBtn.addEventListener("click", () => {
    refreshAchievementDisplay();
    showScreen("home");
  });
  if (ui.resetProgressBtn) {
    ui.resetProgressBtn.addEventListener("click", resetLearningProgress);
  }

  if (ui.openProblemListBtn) {
    ui.openProblemListBtn.addEventListener("click", () => {
      showProblemListScreen();
    });
  }
  if (ui.backHomeFromListBtn) {
    ui.backHomeFromListBtn.addEventListener("click", () => {
      showScreen("home");
      refreshAchievementDisplay();
    });
  }
  if (ui.backHomeFromListTopBtn) {
    ui.backHomeFromListTopBtn.addEventListener("click", () => {
      showScreen("home");
      refreshAchievementDisplay();
    });
  }
  ui.listGradeButtons.forEach((button) => {
    button.addEventListener("click", () => {
      appState.listView.grade = Number(button.dataset.grade);
      updateProblemListGradeButtonState();
      renderProblemList();
    });
  });
  ui.listLangRadios.forEach((radio) => {
    radio.addEventListener("change", () => {
      if (radio.checked) {
        appState.listView.langMode = radio.value;
        renderProblemList();
      }
    });
  });

  ui.wordBank.addEventListener("dragover", (event) => {
    event.preventDefault();
    ui.wordBank.classList.add("drag-over");
  });
  ui.wordBank.addEventListener("dragleave", () => ui.wordBank.classList.remove("drag-over"));
  ui.wordBank.addEventListener("drop", (event) => {
    event.preventDefault();
    ui.wordBank.classList.remove("drag-over");
    if (isAnswerLocked()) {
      return;
    }
    if (!draggingToken || draggingToken.source !== "answer") {
      return;
    }
    saveSnapshot();
    const token = removeTokenFromAnswer(draggingToken.id);
    if (token) {
      appState.wordBank.push(token);
      renderWordAreas();
    }
  });

  ui.answerZone.addEventListener("dragover", (event) => {
    event.preventDefault();
    ui.answerZone.classList.add("drag-over");
    activateNearestGap(event.clientX, event.clientY);
  });
  ui.answerZone.addEventListener("dragleave", () => {
    ui.answerZone.classList.remove("drag-over");
    clearGapHighlight();
    appState.activeGapIndex = null;
  });
  ui.answerZone.addEventListener("drop", (event) => {
    event.preventDefault();
    ui.answerZone.classList.remove("drag-over");
    if (isAnswerLocked()) {
      return;
    }
    if (!draggingToken) {
      return;
    }
    saveSnapshot();
    const insertIndex = Number.isInteger(appState.activeGapIndex) ? appState.activeGapIndex : appState.answer.length;
    insertTokenToAnswer(draggingToken, insertIndex);
    appState.activeGapIndex = null;
    clearGapHighlight();
    renderWordAreas();
  });
  ui.answerZone.addEventListener("click", (event) => {
    if (isAnswerLocked()) {
      return;
    }
    const target = event.target;
    if (!(target instanceof Element)) {
      return;
    }
    const tokenNode = target.closest(".token");
    if (!tokenNode || tokenNode.dataset.source !== "answer") {
      return;
    }
    const tokenId = tokenNode.dataset.id;
    if (!tokenId) {
      return;
    }
    saveSnapshot();
    const moved = removeTokenFromAnswer(tokenId);
    if (!moved) {
      return;
    }
    appState.wordBank.push(moved);
    renderWordAreas();
  });
}

async function startQuiz(grade) {
  appState.grade = grade;
  appState.currentQuestionIndex = 0;
  appState.correctCount = 0;
  appState.wrongCount = 0;
  appState.wrongQuestionIds = [];
  appState.feedback = "";

  try {
    const pool = await loadQuestionsByGrade(grade);
    const count = Math.min(10, pool.length);
    appState.questionSet = weightedSample(pool, count);
    showScreen("quiz");
    loadQuestion();
  } catch (error) {
    window.alert("問題の読み込み中にエラーが発生しました。");
  }
}

async function startSingleQuestion(grade, questionId) {
  appState.grade = grade;
  appState.currentQuestionIndex = 0;
  appState.correctCount = 0;
  appState.wrongCount = 0;
  appState.wrongQuestionIds = [];
  appState.feedback = "";

  try {
    const pool = await loadQuestionsByGrade(grade);
    const question = pool.find((q) => q.id === questionId);
    if (!question) {
      window.alert("問題の読み込み中にエラーが発生しました。");
      return false;
    }
    appState.questionSet = [question];
    showScreen("quiz");
    loadQuestion();
    return true;
  } catch {
    window.alert("問題の読み込み中にエラーが発生しました。");
    return false;
  }
}

async function loadQuestionsByGrade(grade) {
  if (appState.dataCache.has(grade)) {
    return appState.dataCache.get(grade);
  }
  const embedded = window.QUESTION_BANK && window.QUESTION_BANK[String(grade)];
  if (Array.isArray(embedded) && embedded.length > 0) {
    const normalized = embedded.map((question, index) => normalizeQuestion(question, index));
    appState.dataCache.set(grade, normalized);
    return appState.dataCache.get(grade);
  }
  throw new Error(`questions_grade${grade}.js is not loaded`);
}

function normalizeQuestion(question, index) {
  const sentence = typeof question.sentence === "string" ? question.sentence.trim() : "";
  const posList = Array.isArray(question.pos) ? question.pos : [];

  const normalizedTokens = Array.isArray(question.tokens) && question.tokens.length > 0
    ? question.tokens.map((token, tokenIndex) => normalizeToken(token, tokenIndex, question.id, posList[tokenIndex]))
    : buildTokensFromSentence(sentence, posList, question.id, index);

  return {
    ...question,
    sentence,
    tokens: normalizedTokens,
    sentenceStructure: question.sentenceStructure || null,
  };
}

function normalizeToken(token, tokenIndex, questionId, fallbackPos = "") {
  const text = String(token?.text || "");
  return {
    ...token,
    id: token?.id ? String(token.id) : `${questionId || "q"}_${tokenIndex + 1}`,
    text,
    pos: token?.pos || fallbackPos || "",
  };
}

function buildTokensFromSentence(sentence, posList, questionId, fallbackIndex = 0) {
  const words = tokenizeSentenceForPos(sentence, posList.length);
  return words.map((word, tokenIndex) => {
    return {
      id: `${questionId || `q${fallbackIndex + 1}`}_${tokenIndex + 1}`,
      text: word,
      pos: posList[tokenIndex] || "",
    };
  });
}

function tokenizeSentenceForPos(sentence, expectedCount = 0) {
  if (!sentence) {
    return [];
  }

  const strict = expandCannotForExpectedCount(postProcessSentenceTokens(
    sentence.match(/[A-Za-z]+(?:-[A-Za-z]+)*(?:'[A-Za-z]+)?|\d+(?:[.,]\d+)?|[.,!?;:()]/g) || []
  ), expectedCount);
  if (!expectedCount || strict.length === expectedCount) {
    return strict;
  }

  const loose = expandCannotForExpectedCount(postProcessSentenceTokens(
    sentence
      .replace(/([.,!?;:()])/g, " $1 ")
      .split(/\s+/)
      .map((part) => part.trim())
      .filter(Boolean)
  ), expectedCount);
  if (loose.length === expectedCount) {
    return loose;
  }

  return Math.abs(strict.length - expectedCount) <= Math.abs(loose.length - expectedCount)
    ? strict
    : loose;
}

function postProcessSentenceTokens(tokens) {
  const expanded = [];
  tokens.forEach((token) => {
    const raw = String(token);
    const lower = raw.toLowerCase();

    if (lower.endsWith("n't") && raw.length > 3) {
      expanded.push(raw.slice(0, -3));
      expanded.push("n't");
      return;
    }

    expanded.push(raw);
  });

  const merged = [];
  for (let i = 0; i < expanded.length; i += 1) {
    const current = expanded[i];
    const next = expanded[i + 1];
    const nextNext = expanded[i + 2];
    if (
      current
      && next === "."
      && nextNext
      && /^[A-Za-z]{1,3}$/.test(current)
      && /^[A-Z]/.test(nextNext)
    ) {
      merged.push(current);
      i += 1;
      continue;
    }
    merged.push(current);
  }

  return merged;
}

function expandCannotForExpectedCount(tokens, expectedCount) {
  if (!expectedCount || tokens.length >= expectedCount) {
    return tokens;
  }

  let remaining = expectedCount - tokens.length;
  if (remaining <= 0) {
    return tokens;
  }

  const expanded = [];
  tokens.forEach((token) => {
    const raw = String(token);
    if (remaining > 0 && raw.toLowerCase() === "cannot") {
      expanded.push(raw[0] === "C" ? "Can" : "can");
      expanded.push("not");
      remaining -= 1;
      return;
    }
    expanded.push(raw);
  });
  return expanded;
}

async function preloadAllGrades() {
  await Promise.all(SUPPORTED_GRADES.map(async (grade) => {
    try {
      await loadQuestionsByGrade(grade);
    } catch {
      appState.dataCache.set(grade, []);
    }
  }));
}

function showProblemListScreen() {
  showScreen("list");
  updateProblemListGradeButtonState();
  updateProblemListLangRadioState();
  renderProblemList();
}

function updateProblemListGradeButtonState() {
  ui.listGradeButtons.forEach((button) => {
    const active = Number(button.dataset.grade) === appState.listView.grade;
    button.classList.toggle("primary", active);
  });
}

function updateProblemListLangRadioState() {
  ui.listLangRadios.forEach((radio) => {
    radio.checked = radio.value === appState.listView.langMode;
  });
}

function renderProblemList() {
  const grade = appState.listView.grade;
  const mode = appState.listView.langMode;
  const questions = getQuestionPoolSync(grade)
    .slice()
    .sort((a, b) => getDifficultyValue(a) - getDifficultyValue(b));

  if (ui.problemCount) {
    ui.problemCount.textContent = `${formatGradeLabel(grade)}：${questions.length}問`;
  }
  if (!ui.problemList) {
    return;
  }

  ui.problemList.innerHTML = "";
  questions.forEach((question, index) => {
    const card = document.createElement("article");
    card.className = "question-card";

    const title = document.createElement("h3");
    title.textContent = `Question ${index + 1} (difficulty: ${getDifficultyValue(question)})`;
    card.appendChild(title);

    if (mode === "en" || mode === "both" || mode === "full") {
      const en = document.createElement("p");
      en.textContent = buildEnglishSentence(question);
      card.appendChild(en);
    }
    if (mode === "ja" || mode === "both" || mode === "full") {
      const ja = document.createElement("p");
      ja.textContent = question.japanese || "";
      card.appendChild(ja);
    }
    if (mode === "full") {
      card.appendChild(buildPosCardView(question));
      card.appendChild(buildStructureGuideView(question));
    }

    const row = document.createElement("div");
    row.className = "action-row";
    const button = document.createElement("button");
    button.type = "button";
    button.className = "primary";
    button.textContent = "この問題を解く";
    button.addEventListener("click", () => {
      startSingleQuestion(grade, question.id);
    });
    row.appendChild(button);
    card.appendChild(row);

    ui.problemList.appendChild(card);
  });
}

function buildPosCardView(question) {
  const wrap = document.createElement("div");
  const title = document.createElement("h4");
  title.textContent = "";
  wrap.appendChild(title);

  const row = document.createElement("div");
  row.className = "card-row bank-surface";
  (question.tokens || []).forEach((token) => {
    const item = document.createElement("div");
    const colorClass = POS_COLOR_CLASS[token.pos] || "noun";
    item.className = `token pos-${colorClass}`;
    item.innerHTML = `<strong>${token.text}</strong><small>${getTokenLabel(token)}</small>`;
    row.appendChild(item);
  });
  wrap.appendChild(row);
  return wrap;
}

function buildStructureGuideView(question) {
  const wrap = document.createElement("div");
  const title = document.createElement("h4");
  title.textContent = "文構造ガイド";
  wrap.appendChild(title);

  const lines = getStructureGuideLines(question);
  if (lines.length === 0) {
    const p = document.createElement("p");
    p.textContent = "Structure data not found";
    wrap.appendChild(p);
    return wrap;
  }

  const pre = document.createElement("pre");
  pre.className = "structure-tree";
  pre.textContent = lines.join("\n");
  wrap.appendChild(buildStructureRichView(lines));
  wrap.appendChild(pre);
  wrap.appendChild(buildStructureLegend());

  const rawSentence = getGuideSentence(question);
  if (rawSentence) {
    const sentenceLine = document.createElement("p");
    sentenceLine.className = "structure-original-sentence";
    sentenceLine.textContent = rawSentence;
    wrap.appendChild(sentenceLine);
  }

  return wrap;
}

function buildEnglishSentence(question) {
  if (typeof question.sentence === "string" && question.sentence.trim()) {
    return question.sentence.trim();
  }
  const ss = question.sentenceStructure;
  if (ss && typeof ss.sentence === "string" && ss.sentence.trim()) {
    return ss.sentence.trim();
  }
  return (question.tokens || []).map((t) => t.text).filter(Boolean).join(" ");
}

function loadQuestion() {
  const question = appState.questionSet[appState.currentQuestionIndex];
  appState.currentQuestion = question;
  appState.wordBank = shuffle([...question.tokens]);
  appState.answer = [];
  appState.currentAttemptCount = 0;
  appState.historySnapshots = [];
  appState.wrongPositions = new Set();

  ui.feedback.textContent = "";
  ui.feedback.className = "feedback";
  ui.undoBtn.classList.remove("hidden");
  ui.clearBtn.classList.remove("hidden");
  ui.checkBtn.classList.remove("hidden");
  ui.nextBtn.classList.add("hidden");
  ui.structureCard.classList.add("hidden");
  ui.structureSummary.innerHTML = "";
  ui.structurePanel.innerHTML = "";
  ui.structurePanel.classList.add("hidden");

  ui.questionProgress.textContent = `${appState.currentQuestionIndex + 1} / ${appState.questionSet.length}`;
  ui.scoreMini.textContent = ` ${appState.correctCount}   ${appState.wrongCount}`;
  ui.currentGradeLabel.textContent = `レベル: ${formatGradeLabel(appState.grade)}`;
  ui.gradeAchievement.textContent = buildGradeAchievementText(appState.grade);
  ui.questionTitle.textContent = `${appState.currentQuestionIndex + 1}`;

  renderTranslation();
  renderWordAreas();
}

function renderTranslation() {
  if (!appState.currentQuestion) {
    return;
  }
  ui.jpMeaning.textContent = appState.currentQuestion.japanese || "";
}

function renderWordAreas() {
  updateAnswerLockState();
  renderWordBank();
  renderAnswerZone();
}

function renderWordBank() {
  ui.wordBank.innerHTML = "";
  appState.wordBank.forEach((token) => {
    ui.wordBank.appendChild(createTokenElement(token, "bank", false));
  });
}

function renderAnswerZone() {
  ui.answerZone.innerHTML = "";

  for (let i = 0; i <= appState.answer.length; i += 1) {
    ui.answerZone.appendChild(createGapElement(i));
    if (i < appState.answer.length) {
      const token = appState.answer[i];
      const isWrong = appState.wrongPositions.has(i);
      ui.answerZone.appendChild(createTokenElement(token, "answer", isWrong));
    }
  }
}

function createTokenElement(token, source, isWrong) {
  const item = document.createElement("div");
  item.className = `token pos-${getTokenColorClass(token)}`;
  if (isWrong) {
    item.classList.add("wrong");
  }
  item.draggable = !isAnswerLocked();
  item.dataset.id = token.id;
  item.dataset.source = source;
  item.innerHTML = `<strong>${token.text}</strong><small>${getTokenLabel(token)}</small>`;

  item.addEventListener("dragstart", (event) => {
    if (isAnswerLocked()) {
      event.preventDefault();
      return;
    }
    if (event.dataTransfer) {
      event.dataTransfer.setData("text/plain", token.id);
      event.dataTransfer.effectAllowed = "move";
    }
    item.classList.add("dragging");
    draggingToken = { id: token.id, source };
  });

  item.addEventListener("dragend", () => {
    item.classList.remove("dragging");
    clearGapHighlight();
    draggingToken = null;
  });

  if (source === "bank") {
    item.addEventListener("click", () => {
      if (isAnswerLocked()) {
        return;
      }
      saveSnapshot();
      insertTokenToAnswer({ id: token.id, source: "bank" }, appState.answer.length);
      renderWordAreas();
    });
  }

  return item;
}

function createGapElement(index) {
  const gap = document.createElement("div");
  gap.className = "gap-slot";
  gap.dataset.index = index;

  gap.addEventListener("dragover", (event) => {
    if (isAnswerLocked()) {
      return;
    }
    event.preventDefault();
    clearGapHighlight();
    gap.classList.add("active");
    appState.activeGapIndex = Number(gap.dataset.index);
  });

  gap.addEventListener("dragleave", () => {
    gap.classList.remove("active");
  });

  gap.addEventListener("drop", (event) => {
    if (isAnswerLocked()) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    gap.classList.remove("active");
    if (!draggingToken) {
      return;
    }
    saveSnapshot();

    const insertIndex = Number(gap.dataset.index);
    appState.activeGapIndex = insertIndex;
    insertTokenToAnswer(draggingToken, insertIndex);
    appState.activeGapIndex = null;
    renderWordAreas();
  });

  return gap;
}

function isAnswerLocked() {
  return !ui.nextBtn.classList.contains("hidden");
}

function updateAnswerLockState() {
  const locked = isAnswerLocked();
  ui.wordBank.classList.toggle("locked", locked);
  ui.answerZone.classList.toggle("locked", locked);
}

function activateNearestGap(clientX, clientY) {
  const gaps = [...ui.answerZone.querySelectorAll(".gap-slot")];
  if (gaps.length === 0) {
    return;
  }
  const answerTokens = [...ui.answerZone.querySelectorAll(".token[data-source='answer']")];
  if (answerTokens.length === 0) {
    activateGapByIndex(gaps, 0, false);
    return;
  }
  const rects = answerTokens.map((node) => node.getBoundingClientRect());
  const candidates = buildGapCandidates(rects);

  let best = candidates[0];
  let bestScore = Number.POSITIVE_INFINITY;
  candidates.forEach((candidate) => {
    const dx = Math.abs(clientX - candidate.x);
    const dy = Math.abs(clientY - candidate.y);
    const score = dx + (dy * 1.4);
    if (score < bestScore) {
      bestScore = score;
      best = candidate;
    }
  });

  activateGapByIndex(gaps, best.index, best.visual === "row-end");
}

function activateGapByIndex(gaps, index, showAtPrevRowEnd) {
  clearGapHighlight();
  appState.activeGapIndex = index;

  if (showAtPrevRowEnd && index > 0) {
    const answerTokens = [...ui.answerZone.querySelectorAll(".token[data-source='answer']")];
    const prevToken = answerTokens[index - 1];
    if (prevToken) {
      prevToken.classList.add("gap-end-active");
      return;
    }
  }

  const target = gaps.find((gap) => Number(gap.dataset.index) === index);
  if (target) {
    target.classList.add("active");
  }
}

function buildGapCandidates(tokenRects) {
  const count = tokenRects.length;
  const list = [];
  const first = tokenRects[0];
  const last = tokenRects[count - 1];

  list.push({
    index: 0,
    x: first.left - 10,
    y: first.top + (first.height / 2),
    visual: "gap",
  });

  for (let i = 1; i < count; i += 1) {
    const prev = tokenRects[i - 1];
    const curr = tokenRects[i];
    const wrapped = curr.top > prev.top + 4;
    if (wrapped) {
      // Same insertion index can be targeted from both row edges.
      list.push({
        index: i,
        x: prev.right + 10,
        y: prev.top + (prev.height / 2),
        visual: "row-end",
      });
      list.push({
        index: i,
        x: curr.left - 10,
        y: curr.top + (curr.height / 2),
        visual: "gap",
      });
      continue;
    }

    list.push({
      index: i,
      x: (prev.right + curr.left) / 2,
      y: (prev.top + curr.top + prev.height) / 2,
      visual: "gap",
    });
  }

  list.push({
    index: count,
    x: last.right + 10,
    y: last.top + (last.height / 2),
    visual: "row-end",
  });

  return list;
}

function insertTokenToAnswer(payload, targetIndex) {
  let token = null;
  let insertIndex = targetIndex;

  if (payload.source === "bank") {
    token = removeTokenFromWordBank(payload.id);
  } else {
    const fromIndex = appState.answer.findIndex((item) => item.id === payload.id);
    if (fromIndex === -1) {
      return;
    }
    token = appState.answer[fromIndex];
    appState.answer.splice(fromIndex, 1);
    if (fromIndex < insertIndex) {
      insertIndex -= 1;
    }
  }

  if (!token) {
    return;
  }

  appState.answer.splice(insertIndex, 0, token);
  appState.wrongPositions = new Set();
}

function removeTokenFromWordBank(id) {
  const index = appState.wordBank.findIndex((item) => item.id === id);
  if (index === -1) {
    return null;
  }
  return appState.wordBank.splice(index, 1)[0];
}

function removeTokenFromAnswer(id) {
  const index = appState.answer.findIndex((item) => item.id === id);
  if (index === -1) {
    return null;
  }
  return appState.answer.splice(index, 1)[0];
}

function saveSnapshot() {
  appState.historySnapshots.push({
    answer: [...appState.answer],
    wordBank: [...appState.wordBank],
  });

  if (appState.historySnapshots.length > 30) {
    appState.historySnapshots.shift();
  }
}

function undoMove() {
  const snapshot = appState.historySnapshots.pop();
  if (!snapshot) {
    return;
  }
  appState.answer = snapshot.answer;
  appState.wordBank = snapshot.wordBank;
  appState.wrongPositions = new Set();
  renderWordAreas();
}

function clearAnswer() {
  if (appState.answer.length === 0) {
    return;
  }
  saveSnapshot();
  appState.wordBank = [...appState.wordBank, ...appState.answer];
  appState.answer = [];
  appState.wrongPositions = new Set();
  renderWordAreas();
}

function checkAnswer() {
  if (!ui.nextBtn.classList.contains("hidden")) {
    return;
  }

  const question = appState.currentQuestion;
  const expectedTokens = [...question.tokens];

  if (appState.answer.length !== expectedTokens.length) {
    appState.wrongPositions = new Set();
    ui.feedback.textContent = "";
    ui.feedback.className = "feedback ng";
    playWrongSound();
    return;
  }

  const wrongPositions = [];
  for (let i = 0; i < expectedTokens.length; i += 1) {
    const expectedText = expectedTokens[i].text.toLowerCase();
    const actualText = appState.answer[i].text.toLowerCase();
    if (expectedText !== actualText) {
      wrongPositions.push(i);
    }
  }

  const isFirstAttempt = appState.currentAttemptCount === 0;
  appState.currentAttemptCount += 1;

  if (wrongPositions.length === 0) {
    appState.correctCount += 1;
    ui.feedback.textContent = "\u6b63\u89e3\uff01";
    ui.feedback.className = "feedback ok";
    appState.wrongPositions = new Set();
    updateProgress(question.id, true);
    if (isFirstAttempt) {
      updateFirstTryProgress(question.id, appState.grade, true);
      refreshAchievementDisplay();
      ui.gradeAchievement.textContent = buildGradeAchievementText(appState.grade);
    }
    playCorrectSound();
    ui.checkBtn.classList.add("hidden");
    ui.undoBtn.classList.add("hidden");
    ui.clearBtn.classList.add("hidden");
    renderStructureGuide(question);
    ui.structureCard.classList.remove("hidden");
    ui.nextBtn.classList.remove("hidden");
    ui.scoreMini.textContent = ` ${appState.correctCount}   ${appState.wrongCount}`;
  } else {
    appState.wrongCount += 1;
    appState.wrongQuestionIds.push(question.id);
    appState.wrongPositions = new Set(wrongPositions);
    ui.feedback.textContent = "";
    ui.feedback.className = "feedback ng";
    updateProgress(question.id, false);
    if (isFirstAttempt) {
      updateFirstTryProgress(question.id, appState.grade, false);
      refreshAchievementDisplay();
      ui.gradeAchievement.textContent = buildGradeAchievementText(appState.grade);
    }
    playWrongSound();
    renderAnswerZone();
    ui.scoreMini.textContent = ` ${appState.correctCount}   ${appState.wrongCount}`;
  }
}

function goToNextQuestion() {
  appState.currentQuestionIndex += 1;
  if (appState.currentQuestionIndex >= appState.questionSet.length) {
    showResult();
    return;
  }
  loadQuestion();
}

function showResult() {
  const total = appState.questionSet.length;
  const score = Math.round((appState.correctCount / total) * 100);

  ui.finalScore.textContent = `${score}`;
  ui.finalDetail.textContent = `Correct ${appState.correctCount} / ${total} (Wrong ${appState.wrongCount})`;

  const uniqueWrong = [...new Set(appState.wrongQuestionIds)];
  if (uniqueWrong.length === 0) {
    ui.reviewList.textContent = "";
  } else {
    const labels = uniqueWrong.map((id) => {
      const q = appState.questionSet.find((item) => item.id === id);
      return `<li>${q ? q.japanese : id}</li>`;
    }).join("");
    ui.reviewList.innerHTML = `<p>/p><ul>${labels}</ul>`;
  }

  showScreen("result");
}

function renderStructureGuide(question) {
  ui.structureSummary.innerHTML = "";
  ui.structurePanel.innerHTML = "";
  const lines = getStructureGuideLines(question);
  if (lines.length === 0) {
    const line = document.createElement("p");
    line.textContent = "structure is not set.";
    ui.structureSummary.appendChild(line);
    return;
  }

  const pre = document.createElement("pre");
  pre.className = "structure-tree";
  pre.textContent = lines.join("\n");
  ui.structureSummary.appendChild(buildStructureRichView(lines));
  ui.structureSummary.appendChild(pre);
  ui.structureSummary.appendChild(buildStructureLegend());

  const rawSentence = getGuideSentence(question);
  if (rawSentence) {
    const sentenceLine = document.createElement("p");
    sentenceLine.className = "structure-original-sentence";
    sentenceLine.textContent = rawSentence;
    ui.structureSummary.appendChild(sentenceLine);
  }

  ui.structurePanel.classList.add("hidden");
}

function getStructureGuideLines(question) {
  if (typeof question.structure === "string" && question.structure.trim()) {
    return question.structure
      .split(/\r?\n/)
      .map((line) => line.trimEnd())
      .filter((line, index, arr) => line || index < arr.length - 1);
  }

  const ss = question.sentenceStructure;
  if (!ss) {
    return [];
  }
  return buildStructureTreeLines(ss);
}

function getGuideSentence(question) {
  if (typeof question.sentence === "string" && question.sentence.trim()) {
    return question.sentence.trim();
  }
  const ss = question.sentenceStructure;
  if (ss && typeof ss.sentence === "string" && ss.sentence.trim()) {
    return ss.sentence.trim();
  }
  return "";
}

function buildStructureRichView(lines) {
  const safeLines = Array.isArray(lines) ? lines : [];
  const html = structureTextToHtml(safeLines);
  const box = document.createElement("div");
  box.className = "structure2-wrap";
  box.innerHTML = html;
  return box;
}

function escapeHtml(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function roleToClass(role) {
  const head = String(role || "").charAt(0).toLowerCase();
  if (head === "s" || head === "v" || head === "o" || head === "c") {
    return `role-${head}`;
  }
  return "role-x";
}

function roleToLabel(role) {
  const match = /^([SVOC])(\d?)(\'?)$/.exec(String(role || "")) || [];
  const base = (match[1] || String(role || "").charAt(0) || "").toUpperCase();
  const num = match[2] || "";
  const prime = match[3] || (String(role || "").includes("'") ? "'" : "");

  let jp = "要素";
  if (base === "S") jp = "主語";
  if (base === "V") jp = "動詞";
  if (base === "O") jp = "目的語";
  if (base === "C") jp = "補語";

  return `${base}${num}${prime}(${jp})`;
}

function formatInlinePhrases(text) {
  let out = escapeHtml(text);
  out = out.replace(/&lt;([^&]([\s\S]*?)?)&gt;/g, '<span class="phrase noun">&lt;$1&gt;</span>');
  out = out.replace(/\[([^\]]+)\]/g, '<span class="phrase adj">[$1]</span>');
  out = out.replace(/\(([^\)]+)\)/g, '<span class="phrase adv">($1)</span>');
  return out;
}

function structureTextToHtml(lines) {
  const rolePattern = /\{([^{}]+)\}(O1'|O2'|S'|V'|O'|C'|O1|O2|S|V|O|C)/g;
  const lineHtml = lines.map((line) => {
    rolePattern.lastIndex = 0;
    const safeLine = String(line || "");
    const indentMatch = safeLine.match(/^\s*/);
    const indentSpaces = indentMatch ? indentMatch[0].length : 0;
    const trimmed = safeLine.trim();
    let cursor = 0;
    let inner = "";
    let match;

    while ((match = rolePattern.exec(trimmed)) !== null) {
      if (match.index > cursor) {
        inner += `<span class="plain">${formatInlinePhrases(trimmed.slice(cursor, match.index))}</span>`;
      }
      const content = formatInlinePhrases(match[1]);
      const role = match[2];
      inner += `<span class="svoc ${roleToClass(role)}"><span class="svoc-text">${content}</span><span class="svoc-label">${escapeHtml(roleToLabel(role))}</span></span>`;
      cursor = match.index + match[0].length;
    }

    if (cursor < trimmed.length) {
      inner += `<span class="plain">${formatInlinePhrases(trimmed.slice(cursor))}</span>`;
    }
    if (!inner) {
      inner = `<span class="plain">${formatInlinePhrases(trimmed)}</span>`;
    }
    return `<div class="line" style="padding-left:${indentSpaces * 12}px">${inner}</div>`;
  }).join("");

  return `<div class="structure2">${lineHtml}</div>`;
}

function buildStructureLegend() {
  const help = document.createElement("p");
  help.className = "structure-legend";
  help.textContent = "<> 名詞句・節、[] 形容詞句・節、() 副詞句・節";
  return help;
}

function buildClauseTypeLabel(type, subordinate) {
  if (!subordinate) {
    return "";
  }
  if (type === "noun") {
    return "";
  }
  if (type === "adjective") {
    return "";
  }
  if (type === "adverbial") {
    return "";
  }
  return "";
}

function getClauseWrapMarks(type) {
  if (type === "noun" || type === "noun_clause") {
    return { open: "<", close: ">" };
  }
  if (type === "adjective" || type === "adjective_clause") {
    return { open: "[", close: "]" };
  }
  if (type === "adverbial" || type === "adverbial_clause") {
    return { open: "(", close: ")" };
  }
  return { open: "", close: "" };
}

function buildStructureTreeLines(sentenceStructure) {
  if (Array.isArray(sentenceStructure.treeLinesJa) && sentenceStructure.treeLinesJa.length > 0) {
    return sentenceStructure.treeLinesJa;
  }
  if (Array.isArray(sentenceStructure.treeLines) && sentenceStructure.treeLines.length > 0) {
    return sentenceStructure.treeLines;
  }

  const clauses = Array.isArray(sentenceStructure.clauses) ? sentenceStructure.clauses : [];
  const mainClauses = clauses
    .filter((c) => c.type === "main")
    .sort((a, b) => a.span[0] - b.span[0]);
  const subordinateClauses = clauses
    .filter((c) => c.type !== "main")
    .sort((a, b) => a.span[0] - b.span[0]);
  const lines = [];

  if (mainClauses.length === 0) {
    if (subordinateClauses.length > 0) {
      subordinateClauses.forEach((clause, idx) => {
        lines.push(formatClauseHeadLine(clause, true, true));
        if (idx < subordinateClauses.length - 1) {
          lines.push("");
        }
      });
      return lines;
    }
    return ["?????????"];
  }

  if (mainClauses.length === 1) {
    lines.push(...formatMainClauseAsTree(mainClauses[0]));
  } else {
    mainClauses.forEach((clause, index) => {
      const ownMarker = clause.marker ? String(clause.marker).toLowerCase() : "";
      const nextClause = index < mainClauses.length - 1 ? mainClauses[index + 1] : null;
      const nextMarker = nextClause && nextClause.marker ? String(nextClause.marker).toLowerCase() : "";

      if (index > 0 && clause.marker) {
        lines.push(formatTreeElement({ text: String(clause.marker), stanza: "cc" }, `X${index}`));
      }

      const excludeTexts = [ownMarker, nextMarker].filter(Boolean);
      lines.push(...formatMainClauseAsStandaloneTree(clause, index + 1, { excludeTexts }));

      if (index < mainClauses.length - 1) {
        lines.push("");
      }
    });
  }

  if (subordinateClauses.length > 0) {
    lines.push("");
    subordinateClauses.forEach((clause, idx) => {
      lines.push(formatClauseHeadLine(clause, true, true));
      if (idx < subordinateClauses.length - 1) {
        lines.push("");
      }
    });
  }

  return lines;
}

function formatClauseHeadLine(clause, subordinate, addComma) {
  const parts = [];
  const prime = subordinate ? "'" : "";
  const marker = clause.marker ? `${capitalizeWord(clause.marker)} ` : "";
  const elements = clause.elements || {};
  if (marker) {
    parts.push(marker.trim());
  }
  ["S", "V", "O", "C", "M", "A", "X"].forEach((key) => {
    const value = elements[key];
    if (value === undefined) {
      return;
    }
    if (Array.isArray(value)) {
      value.forEach((item, idx) => {
        if (key === "M" || key === "A") {
          parts.push(formatTreeModifierHead(item, idx + 1, prime));
          return;
        }
        parts.push(formatTreeElement(item, `${key}${idx + 1}${prime}`));
      });
      return;
    }
    if (key === "M" || key === "A") {
      parts.push(formatTreeModifierHead(value, 1, prime));
      return;
    }
    parts.push(formatTreeElement(value, `${key}${prime}`));
  });
  const wrap = getClauseWrapMarks(clause.type || "main");
  const body = parts.join(" ");
  const punct = addComma ? "," : "";
  return `${wrap.open}${body}${wrap.close}${punct}`;
}

function formatMainClauseAsTree(clause) {
  const lines = [];
  const elements = clause.elements || {};
  const verbs = toElementArray(elements.V);
  const subjects = toElementArray(elements.S);

  const rootSubject = subjects[0];
  if (rootSubject !== undefined && rootSubject !== null) {
    lines.push(formatTreeElement(rootSubject, "S"));
  }

  if (verbs.length === 0) {
    return lines;
  }

  const objects = toElementArray(elements.O);
  const complements = toElementArray(elements.C);
  const modifiers = toElementArray(elements.M !== undefined ? elements.M : elements.A);
  const extras = toElementArray(elements.X);

  verbs.forEach((verb, index) => {
    const isLastVerb = index === verbs.length - 1;
    const verbPrefix = isLastVerb ? " " : " ";
    lines.push(`${verbPrefix}${formatTreeElement(verb, `V${index + 1}`)}`);

    const children = [];
    if (subjects.length === verbs.length && index > 0) {
      const branchSubject = subjects[index];
      const showBranchSubject = branchSubject !== undefined
        && branchSubject !== null
        && getElementText(branchSubject) !== getElementText(rootSubject);
      if (showBranchSubject) {
        children.push(formatTreeElement(branchSubject, `S${index + 1}`));
      }
    }
    if (objects[index] !== undefined && objects[index] !== null) {
      children.push(formatTreeElement(objects[index], `O${index + 1}`));
    }
    if (complements[index] !== undefined && complements[index] !== null) {
      children.push(formatTreeElement(complements[index], `C${index + 1}`));
    }
    const modifierItems = getItemsForVerbIndex(modifiers, index, verbs.length);
    modifierItems.filter((item) => item !== undefined && item !== null).forEach((item, modIndex) => {
      children.push(formatTreeAdverbial(item, modIndex + 1));
    });
    if (extras[index] !== undefined && extras[index] !== null) {
      children.push(formatTreeElement(extras[index], `X${index + 1}`));
    }

    children.forEach((child, childIndex) => {
      const base = isLastVerb ? "    " : "  ";
      const branch = childIndex === children.length - 1 ? " " : " ";
      lines.push(`${base}${branch}${child}`);
    });
  });

  return lines;
}

function formatMainClauseAsStandaloneTree(clause, slotIndex, options = {}) {
  const lines = [];
  const elements = clause.elements || {};
  const excludeSet = new Set((options.excludeTexts || []).map((x) => String(x || "").trim().toLowerCase()).filter(Boolean));

  const subject = toElementArray(elements.S).find((item) => item !== undefined && item !== null);
  const verb = toElementArray(elements.V).find((item) => item !== undefined && item !== null);
  const object = toElementArray(elements.O).find((item) => item !== undefined && item !== null);
  const complement = toElementArray(elements.C).find((item) => item !== undefined && item !== null);
  const extras = toElementArray(elements.X)
    .filter((item) => item !== undefined && item !== null)
    .filter((item) => {
      const text = getElementText(item).toLowerCase();
      return text && !excludeSet.has(text);
    });
  const modifiers = toElementArray(elements.M !== undefined ? elements.M : elements.A)
    .filter((item) => item !== undefined && item !== null);

  if (subject) {
    lines.push(formatTreeElement(subject, `S${slotIndex}`));
  }
  if (!verb) {
    return lines;
  }

  lines.push(`\u2514\u2500 ${formatTreeElement(verb, `V${slotIndex}`)}`);

  const children = [];
  if (object) {
    children.push(formatTreeElement(object, `O${slotIndex}`));
  }
  if (complement) {
    children.push(formatTreeElement(complement, `C${slotIndex}`));
  }
  modifiers.forEach((item, idx) => {
    children.push(formatTreeAdverbial(item, idx + 1));
  });
  extras.forEach((item, idx) => {
    children.push(formatTreeElement(item, `X${idx + 1}`));
  });

  children.forEach((child, idx) => {
    const branch = idx === children.length - 1 ? "\u2514\u2500 " : "\u251c\u2500 ";
    lines.push(`    ${branch}${child}`);
  });

  return lines;
}

function getItemsForVerbIndex(items, verbIndex, totalVerbs) {
  if (items.length === 0) {
    return [];
  }
  const defined = items.filter((item) => item !== undefined && item !== null);
  if (totalVerbs <= 1) {
    return defined;
  }
  if (items.length === totalVerbs) {
    return items[verbIndex] !== undefined && items[verbIndex] !== null ? [items[verbIndex]] : [];
  }
  if (verbIndex < totalVerbs - 1) {
    return items[verbIndex] !== undefined && items[verbIndex] !== null ? [items[verbIndex]] : [];
  }
  return items.slice(verbIndex).filter((item) => item !== undefined && item !== null);
}

function formatTreeElement(value, roleLabel) {
  if (Array.isArray(value)) {
    const merged = value.map((item) => getElementText(item)).filter(Boolean).join(" / ");
    return `[${merged}]${roleLabel}`;
  }
  if (value && typeof value === "object") {
    const type = value.type || "";
    const text = String(value.text || "");
    const stanza = buildStanzaSuffix(value.stanza, roleLabel);
    if (type === "noun_clause") {
      return `<${text}>${roleLabel}${stanza}`;
    }
    return `[${text}]${roleLabel}${stanza}`;
  }
  return `[${String(value)}]${roleLabel}`;
}

function formatTreeAdverbial(value, index = 1) {
  if (value && typeof value === "object" && !Array.isArray(value)) {
    const stanza = buildStanzaSuffix(value.stanza, "M");
    return `(${String(value.text || "")})M${index}${stanza}`;
  }
  return `(${String(value)})M${index}`;
}

function formatTreeModifierHead(value, index = 1, prime = "") {
  if (value && typeof value === "object" && !Array.isArray(value)) {
    const stanza = buildStanzaSuffix(value.stanza, "M");
    return `(${String(value.text || "")})M${index}${prime}${stanza}`;
  }
  return `(${String(value)})M${index}${prime}`;
}

function buildStanzaSuffix(raw, roleLabel = "") {
  if (!raw) {
    return "";
  }
  let source = String(raw);
  if (String(roleLabel).startsWith("V") && source === "advcl") {
    source = "verb_subordinate";
  }
  if (String(roleLabel).startsWith("V") && source === "conj") {
    source = "root";
  }
  const label = STANZA_LABEL_JA[source] || source;
  return `(${label})`;
}

function toElementArray(value) {
  if (value === undefined) {
    return [];
  }
  return Array.isArray(value) ? value : [value];
}

function getElementText(value) {
  if (value === null || value === undefined) {
    return "";
  }
  if (typeof value === "object" && !Array.isArray(value)) {
    return String(value.text || "").trim();
  }
  return String(value).trim();
}

function capitalizeWord(word) {
  if (!word) {
    return "";
  }
  const text = String(word);
  return text.charAt(0).toUpperCase() + text.slice(1);
}

function getTokenLabel(token) {
  if (token.pos === "PARTICLE") {
    return "";
  }
  return POS_LABEL[token.pos] || token.pos;
}

function getTokenColorClass(token) {
  if (token.pos === "PARTICLE") {
    return "verb";
  }
  return POS_COLOR_CLASS[token.pos] || "noun";
}

function quitQuizToHome() {
  const ok = window.confirm("トップへ戻ります。よろしいですか？");
  if (!ok) {
    return;
  }
  refreshAchievementDisplay();
  showScreen("home");
}

function showScreen(key) {
  Object.values(ui.screens).forEach((element) => element.classList.remove("active"));
  ui.screens[key].classList.add("active");
}

function weightedSample(questions, count) {
  const unseenTarget = Math.min(count, Math.max(0, Math.ceil(count * 0.4)));
  const unseenPool = [];
  const seenPool = [];

  questions.forEach((question) => {
    const seen = hasSeenQuestion(question.id);
    const entry = { question, weight: Math.max(0.2, getQuestionWeight(question.id)) };
    if (seen) {
      seenPool.push(entry);
    } else {
      unseenPool.push(entry);
    }
  });

  const picked = [];
  pickLowDifficultyInto(picked, unseenPool, unseenTarget);
  pickWeightedInto(picked, seenPool, count - picked.length);

  if (picked.length < count) {
    pickLowDifficultyInto(picked, unseenPool, count - picked.length);
  }

  if (picked.length < count) {
    pickWeightedInto(picked, seenPool, count - picked.length);
  }

  return picked;
}

function pickLowDifficultyInto(targetQuestions, pool, takeCount) {
  if (takeCount <= 0 || pool.length === 0) {
    return;
  }
  pool.sort((a, b) => getDifficultyValue(a.question) - getDifficultyValue(b.question));
  const take = Math.min(takeCount, pool.length);
  for (let i = 0; i < take; i += 1) {
    targetQuestions.push(pool[i].question);
  }
  pool.splice(0, take);
}

function getDifficultyValue(question) {
  const n = Number(question.difficulty);
  return Number.isFinite(n) ? n : 999;
}

function pickWeightedInto(targetQuestions, pool, takeCount) {
  const goalLength = targetQuestions.length + takeCount;
  while (targetQuestions.length < goalLength && pool.length > 0) {
    const totalWeight = pool.reduce((sum, item) => sum + item.weight, 0);
    let random = Math.random() * totalWeight;
    let chosenIndex = pool.length - 1;

    for (let i = 0; i < pool.length; i += 1) {
      random -= pool[i].weight;
      if (random <= 0) {
        chosenIndex = i;
        break;
      }
    }

    targetQuestions.push(pool[chosenIndex].question);
    pool.splice(chosenIndex, 1);
  }
}

function hasSeenQuestion(questionId) {
  const progress = loadProgress();
  const stats = progress[questionId];
  if (!stats) {
    return false;
  }
  return (stats.correctCount || 0) + (stats.wrongCount || 0) > 0;
}

function getQuestionWeight(questionId) {
  const progress = loadProgress();
  const stats = progress[questionId];
  if (!stats) {
    return 2.4;
  }

  const seen = stats.correctCount + stats.wrongCount;
  const wrongRate = seen === 0 ? 0 : stats.wrongCount / seen;
  const now = Date.now();
  const lastSeen = Number(stats.lastSeenAt || 0);
  const recentPenalty = (now - lastSeen) < (24 * 60 * 60 * 1000) ? 0.55 : 1;

  return (1 + wrongRate * 2 + (stats.recentWrongStreak || 0) * 0.5) * recentPenalty;
}

function updateProgress(questionId, correct) {
  const progress = loadProgress();
  const current = progress[questionId] || {
    correctCount: 0,
    wrongCount: 0,
    lastSeenAt: 0,
    recentWrongStreak: 0,
  };

  if (correct) {
    current.correctCount += 1;
    current.recentWrongStreak = 0;
  } else {
    current.wrongCount += 1;
    current.recentWrongStreak += 1;
  }
  current.lastSeenAt = Date.now();

  progress[questionId] = current;
  localStorage.setItem(STORAGE_KEYS.progress, JSON.stringify(progress));
}

function updateFirstTryProgress(questionId, grade, firstTryCorrect) {
  const progress = loadProgress();
  const current = progress[questionId] || {
    correctCount: 0,
    wrongCount: 0,
    lastSeenAt: 0,
    recentWrongStreak: 0,
  };
  current.grade = grade;
  current.lastFirstTryCorrect = firstTryCorrect;
  progress[questionId] = current;
  localStorage.setItem(STORAGE_KEYS.progress, JSON.stringify(progress));
}

function loadProgress() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEYS.progress)) || {};
  } catch {
    return {};
  }
}

function buildGradeAchievementText(grade) {
  const summary = getGradeAchievement(grade);
  return `Rate: ${summary.rate}% (${summary.firstTryCorrect}/${summary.total})`;
}

function getGradeAchievement(grade) {
  const questionList = getQuestionPoolSync(grade);
  const total = questionList.length;
  if (total === 0) {
    return { firstTryCorrect: 0, total: 0, rate: 0 };
  }

  const progress = loadProgress();
  const firstTryCorrect = questionList.reduce((sum, q) => {
    return sum + (progress[q.id]?.lastFirstTryCorrect ? 1 : 0);
  }, 0);
  const rate = Math.round((firstTryCorrect / total) * 100);
  return { firstTryCorrect, total, rate };
}

function getQuestionPoolSync(grade) {
  if (!grade) {
    return [];
  }
  if (appState.dataCache.has(grade)) {
    return appState.dataCache.get(grade);
  }
  return [];
}

function refreshAchievementDisplay() {
  SUPPORTED_GRADES.forEach((grade) => {
    const node = ui.homeRateElements[grade];
    if (!node) {
      return;
    }
    const summary = getGradeAchievement(grade);
    node.textContent = `Rate: ${summary.rate}% (${summary.firstTryCorrect}/${summary.total})`;
  });
}

function formatGradeLabel(grade) {
  const num = Number(grade);
  if (GRADE_LABELS[num]) {
    return GRADE_LABELS[num];
  }
  return `Grade ${grade}`;
}

function resetLearningProgress() {
  const ok = window.confirm(
    "学習記録をリセットします。よろしいですか？"
  );
  if (!ok) {
    return;
  }
  localStorage.removeItem(STORAGE_KEYS.progress);
  refreshAchievementDisplay();
  if (appState.grade) {
    ui.gradeAchievement.textContent = buildGradeAchievementText(appState.grade);
  }
  window.alert("学習記録をリセットしました。");
}

function loadSettings() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEYS.settings)) || {};
    appState.settings.soundOn = saved.soundOn !== false;
  } catch {
    appState.settings.soundOn = true;
  }

  ui.soundToggle.checked = appState.settings.soundOn;
}

function saveSettings() {
  localStorage.setItem(STORAGE_KEYS.settings, JSON.stringify(appState.settings));
}

function clearGapHighlight() {
  document.querySelectorAll(".gap-slot.active").forEach((node) => node.classList.remove("active"));
  document.querySelectorAll(".token.gap-end-active").forEach((node) => node.classList.remove("gap-end-active"));
}

function shuffle(items) {
  const array = [...items];
  for (let i = array.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
}

function playCorrectSound() {
  if (!appState.settings.soundOn) {
    return;
  }
  ensureAudio();
  playTone(784, 0.09, "sine", 0);
  playTone(988, 0.14, "sine", 0.1);
}

function playWrongSound() {
  if (!appState.settings.soundOn) {
    return;
  }
  ensureAudio();
  playTone(180, 0.2, "square", 0);
}

function ensureAudio() {
  if (!audioCtx) {
    audioCtx = new (window.AudioContext || window.webkitAudioContext)();
  }
  if (audioCtx.state === "suspended") {
    audioCtx.resume();
  }
}

function playTone(freq, duration, type, delaySeconds) {
  const start = audioCtx.currentTime + delaySeconds;
  const oscillator = audioCtx.createOscillator();
  const gainNode = audioCtx.createGain();

  oscillator.type = type;
  oscillator.frequency.setValueAtTime(freq, start);

  gainNode.gain.setValueAtTime(0, start);
  gainNode.gain.linearRampToValueAtTime(0.2, start + 0.01);
  gainNode.gain.linearRampToValueAtTime(0, start + duration);

  oscillator.connect(gainNode);
  gainNode.connect(audioCtx.destination);
  oscillator.start(start);
  oscillator.stop(start + duration + 0.03);
}

