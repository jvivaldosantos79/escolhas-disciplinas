const DATA_SOURCE = "alunos.csv";
const CURRENT_PROCESS_ID = "10_escolhas";
const processesConfig = {
  "12_opcionais": {
    year: "12",
    title: "Escolhas de disciplinas do 12.º ano",
    mode: "12-opcionais",
    storageKey: "escolhas12ano.resultados",
    downloadName: "escolhas-12ano.csv"
  },
  "10_escolhas": {
    year: "10",
    title: "Escolha de curso do 10.º ano",
    mode: "10-curso",
    storageKey: "escolhas10ano.resultados",
    downloadName: "escolhas-10ano.csv"
  }
};
const CURRENT_PROCESS = processesConfig[CURRENT_PROCESS_ID] || processesConfig["12_opcionais"];
const CURRENT_PROCESS_YEAR = CURRENT_PROCESS.year;
const STORAGE_KEY = CURRENT_PROCESS.storageKey;
const SUPABASE_URL = "https://rygyxkcgvimvommdnuiw.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_RVB9XvY8C7qLzbxvGc7E-A_ch_FCYH6";
const SUPABASE_MODULE_URLS = [
  "https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2/+esm",
  "https://esm.sh/@supabase/supabase-js@2"
];

// Preenche este valor no Azure/Microsoft Entra depois de registares a aplicação SPA.
const MSAL_CLIENT_ID = "585048a3-1008-413b-af8a-cabb2ca780ca";
const MSAL_TENANT = "57b6e55f-5343-418f-b407-7ccfac24fab0";
const MSAL_REDIRECT_URI = window.location.origin + window.location.pathname;
const HOME_URL = window.location.origin + window.location.pathname;
const MSAL_SCRIPT_URLS = [
  "https://alcdn.msauth.net/browser/2.38.2/js/msal-browser.min.js",
  "https://alcdn.msauth.net/browser/2.14.2/js/msal-browser.min.js",
  "https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.38.3/dist/msal-browser.min.js",
  "https://unpkg.com/@azure/msal-browser@2.38.3/dist/msal-browser.min.js"
];

const courseRules = {
  CienciasTecnologias: {
    label: "Ciências e Tecnologias",
    requiredGroupName: "opções específicas d",
    required: ["Biologia", "Física", "Química"],
    optional: [
      "Aplicações Informáticas B",
      "Direito",
      "Economia C",
      "Filosofia A",
      "Geografia C",
      "Inglês - Língua Estrangeira I, II ou III",
      "Psicologia B"
    ],
    ruleText: "Deves indicar 3 prioridades. Em cada prioridade, escolhe duas disciplinas anuais; pelo menos uma tem de ser Biologia, Física ou Química."
  },
  Socioeconomicas: {
    label: "Ciências Socioeconómicas",
    requiredGroupName: "opções específicas d",
    required: ["Economia C", "Geografia C", "Sociologia"],
    optional: [
      "Aplicações Informáticas B",
      "Direito",
      "Filosofia A",
      "Inglês - Língua Estrangeira I, II ou III",
      "Psicologia B"
    ],
    ruleText: "Deves indicar 3 prioridades. Em cada prioridade, escolhe duas disciplinas anuais; pelo menos uma tem de ser Economia C, Geografia C ou Sociologia."
  },
  CienciasTecnologiasCambridge: {
    label: "Ciências e Tecnologias Cambridge",
    cambridge: true,
    cambridgeType: "ct",
    requiredGroupName: "opções específicas d",
    required: ["Biologia", "Física", "Química"],
    optional: [],
    ruleText: "Indica se vais fazer exame de Maths em outubro. Escolhe ou confirma Psychology e Economics e seleciona uma disciplina adicional entre Biologia, Física ou Química."
  },
  SocioeconomicasCambridge: {
    label: "Ciências Socioeconómicas Cambridge",
    cambridge: true,
    cambridgeType: "se",
    requiredGroupName: "opções específicas d",
    required: ["Geografia C", "Sociologia", "Economia C"],
    optional: [],
    ruleText: "Indica se vais fazer exame de Maths em outubro. Se sim, escolhe Psychology, Economics ou ambas. Dependendo da escolha, pode ser necessária uma disciplina adicional."
  },
  LinguasHumanidades: {
    label: "Línguas e Humanidades",
    requiredGroupName: "opções específicas d",
    required: [
      "Filosofia A",
      "Geografia C",
      "Inglês - Língua Estrangeira I, II ou III",
      "Psicologia B",
      "Sociologia"
    ],
    optional: ["Aplicações Informáticas B", "Direito", "Economia C"],
    ruleText: "Deves indicar 3 prioridades. Em cada prioridade, escolhe duas disciplinas anuais; pelo menos uma tem de ser Filosofia A, Geografia C, Inglês, Psicologia B ou Sociologia."
  },
  ArtesVisuais: {
    label: "Artes Visuais",
    requiredGroupName: "opções específicas d",
    required: ["Oficina de Artes", "Oficina de Multimédia B"],
    optional: [
      "Aplicações Informáticas B",
      "Direito",
      "Economia C",
      "Filosofia A",
      "Geografia C",
      "Inglês - Língua Estrangeira I, II ou III",
      "Psicologia B"
    ],
    ruleText: "Deves indicar 3 prioridades. Em cada prioridade, escolhe duas disciplinas anuais; pelo menos uma tem de ser Oficina de Artes ou Oficina de Multimédia B. Se escolheres apenas uma opção d), podes completar com uma opção e)."
  }
};

const courseAliases = {
  "CCH - Ciências e Tecnologias Cambridge": "CienciasTecnologiasCambridge",
  "CCH - Ciências Socioeconómicas Cambridge": "SocioeconomicasCambridge",
  "CCH - Ciências e Tecnologias": "CienciasTecnologias",
  "CCH - Ciências Socioeconómicas": "Socioeconomicas",
  "CCH - Línguas e Humanidades": "LinguasHumanidades",
  "CCH - Artes Visuais": "ArtesVisuais",
  CienciasTecnologias: "CienciasTecnologias",
  Socioeconomicas: "Socioeconomicas",
  LinguasHumanidades: "LinguasHumanidades",
  ArtesVisuais: "ArtesVisuais"
};

const tenthGradeCourses = {
  CienciasTecnologias: {
    label: "Ciências e Tecnologias",
    cambridge: false,
    automaticSubjects: ["Física e Química A"],
    optionSubjects: ["Biologia e Geologia", "Geometria Descritiva A"],
    ruleText: "Ficas inscrito em Física e Química A e escolhes uma disciplina entre Biologia e Geologia e Geometria Descritiva A."
  },
  CienciasTecnologiasCambridge: {
    label: "Ciências e Tecnologias Cambridge",
    cambridge: true,
    automaticSubjects: ["Física e Química A"],
    optionSubjects: ["Biologia e Geologia", "Geometria Descritiva A"],
    ruleText: "Ficas inscrito em Física e Química A e escolhes uma disciplina entre Biologia e Geologia e Geometria Descritiva A. O percurso Cambridge fica sujeito a validação pelo Colégio, quando aplicável."
  },
  Socioeconomicas: {
    label: "Ciências Socioeconómicas",
    cambridge: false,
    automaticSubjects: ["Economia A", "Geografia A"],
    optionSubjects: [],
    ruleText: "Ficas inscrito automaticamente em Economia A e Geografia A."
  },
  SocioeconomicasCambridge: {
    label: "Ciências Socioeconómicas Cambridge",
    cambridge: true,
    automaticSubjects: ["Economia A", "Geografia A"],
    optionSubjects: [],
    ruleText: "Ficas inscrito automaticamente em Economia A e Geografia A. O percurso Cambridge fica sujeito a validação pelo Colégio, quando aplicável."
  },
  LinguasHumanidades: {
    label: "Línguas e Humanidades",
    cambridge: false,
    automaticSubjects: ["Geografia A", "MACS"],
    optionSubjects: [],
    ruleText: "Ficas inscrito automaticamente em Geografia A e MACS."
  },
  ArtesVisuais: {
    label: "Artes Visuais",
    cambridge: false,
    automaticSubjects: ["Desenho A", "História e Cultura das Artes", "Geometria Descritiva A"],
    optionSubjects: [],
    ruleText: "Ficas inscrito automaticamente em Desenho A, História e Cultura das Artes e Geometria Descritiva A."
  }
};

const loginMessage = document.querySelector("#login-message");
const pageTitle = document.querySelector("#page-title");
const authWarning = document.querySelector("#auth-warning");
const signInButton = document.querySelector("#sign-in");
const headerSession = document.querySelector("#header-session");
const headerAuthEmail = document.querySelector("#header-auth-email");
const profilePhoto = document.querySelector("#profile-photo");
const headerSignOutButton = document.querySelector("#header-sign-out");
const homeLink = document.querySelector("#home-link");
const faqNav = document.querySelector("#faq-nav");
const adminNav = document.querySelector("#admin-nav");
const studentArea = document.querySelector("#student-area");
const summaryArea = document.querySelector("#summary-area");
const faqArea = document.querySelector("#faq-area");
const summaryMessage = document.querySelector("#summary-message");
const summaryTable = document.querySelector("#summary-table");
const goHomeButton = document.querySelector("#go-home");
const editChoiceButton = document.querySelector("#edit-choice");
const faqBackButton = document.querySelector("#faq-back");
const choiceStatus = document.querySelector("#choice-status");
const choicesForm = document.querySelector("#choices-form");
const standardChoices = document.querySelector("#standard-choices");
const cambridgeChoices = document.querySelector("#cambridge-choices");
const tenthChoices = document.querySelector("#tenth-choices");
const prioritiesList = document.querySelector("#priorities-list");
const tenthCourseOptions = document.querySelector("#tenth-course-options");
const tenthAutomaticSubjects = document.querySelector("#tenth-automatic-subjects");
const tenthOptionSubjects = document.querySelector("#tenth-option-subjects");
const tenthOptionList = document.querySelector("#tenth-option-list");
const tenthCambridgeNotice = document.querySelector("#tenth-cambridge-notice");
const cambridgeMainOptions = document.querySelector("#cambridge-main-options");
const cambridgeExtraOptions = document.querySelector("#cambridge-extra-options");
const cambridgeExtraNote = document.querySelector("#cambridge-extra-note");
const cambridgeExtraList = document.querySelector("#cambridge-extra-list");
const cambridgeAutoMessage = document.querySelector("#cambridge-auto-message");
const validationMessage = document.querySelector("#validation-message");
const submitChoiceButton = document.querySelector("#submit-choice");
const cancelEditButton = document.querySelector("#cancel-edit");
const confirmation = document.querySelector("#confirmation");
const csvOutput = document.querySelector("#csv-output");
const adminDashboard = document.querySelector("#admin-dashboard");
const adminStatsDashboard = document.querySelector("#admin-stats-dashboard");
const adminResults = document.querySelector("#admin-results");
const refreshResultsButton = document.querySelector("#refresh-results");
const exportCsvButton = document.querySelector("#export-csv");
const downloadCsvButton = document.querySelector("#download-csv");
const clearResultsButton = document.querySelector("#clear-results");
const openAdminStatsButton = document.querySelector("#open-admin-stats");
const openAdminToolsButton = document.querySelector("#open-admin-tools");
const backAdminFromStatsButton = document.querySelector("#back-admin-from-stats");
const backAdminButton = document.querySelector("#back-admin");
const adminFiltersForm = document.querySelector("#admin-filters");
const filterCourse = document.querySelector("#filter-course");
const filterClass = document.querySelector("#filter-class");
const filterStatus = document.querySelector("#filter-status");
const filterSubmission = document.querySelector("#filter-submission");
const filterSearch = document.querySelector("#filter-search");
const clearFiltersButton = document.querySelector("#clear-filters");

let students = [];
let currentStudent = null;
let currentChoice = null;
let msalClient = null;
let signedInAccount = null;
let authInitPromise = null;
let supabaseClient = null;
let supabaseInitPromise = null;
let isAdminUser = false;
let adminStudentsCache = [];
let adminChoicesCache = [];
let adminStudentStatusesCache = new Map();
let lastStudentView = "home";

const loginRequest = {
  scopes: ["openid", "profile", "email", "User.Read"]
};

const studentRepository = {
  async getAll() {
    const client = await ensureSupabaseReady();

    if (client) {
      const { data, error } = await client
        .from("alunos")
        .select("aluno_id,nome,turma,curso,email,processo_id,ano")
        .eq("processo_id", CURRENT_PROCESS_ID)
        .order("turma", { ascending: true })
        .order("nome", { ascending: true });

      if (!error && data) {
        return data.map(normalizeStudent);
      }
    }

    const response = await fetch(DATA_SOURCE);

    if (!response.ok) {
      throw new Error("Não foi possível carregar a base de dados de alunos.");
    }

    const text = await response.text();
    return parseCsv(text);
  },

  async findById(studentId) {
    const client = await ensureSupabaseReady();

    if (client) {
      const { data, error } = await client
        .from("alunos")
        .select("aluno_id,nome,turma,curso,email,processo_id,ano")
        .eq("aluno_id", studentId)
        .eq("processo_id", CURRENT_PROCESS_ID)
        .maybeSingle();

      if (error) {
        throw new Error("Não foi possível consultar a base de dados de alunos.");
      }

      return data ? normalizeStudent(data) : null;
    }

    if (students.length === 0) {
      students = await this.getAll();
    }

    return students.find((student) => student.aluno_id === studentId && getProcessId(student) === CURRENT_PROCESS_ID);
  },

  async findByEmail(email) {
    const normalizedEmail = email.toLowerCase();
    const client = await ensureSupabaseReady();

    if (client) {
      const { data, error } = await client
        .from("alunos")
        .select("aluno_id,nome,turma,curso,email,processo_id,ano")
        .eq("email", normalizedEmail)
        .eq("processo_id", CURRENT_PROCESS_ID)
        .maybeSingle();

      if (error) {
        throw new Error("Não foi possível validar os dados do aluno.");
      }

      return data ? normalizeStudent(data) : null;
    }

    if (students.length === 0) {
      students = await this.getAll();
    }

    return students.find((student) => (student.email || "").toLowerCase() === normalizedEmail && getProcessId(student) === CURRENT_PROCESS_ID) || null;
  }
};

const choiceRepository = {
  async getByAlunoId(alunoId) {
    const client = await ensureSupabaseReady();

    if (client) {
      const { data, error } = await client
        .from("escolhas")
        .select("*")
        .eq("aluno_id", String(alunoId))
        .eq("processo_id", CURRENT_PROCESS_ID)
        .maybeSingle();

      if (error) {
        throw new Error("Não foi possível consultar a submissão existente.");
      }

      return data ? mapChoiceFromDatabase(data) : null;
    }

    const choices = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
    return choices.find((choice) => choice.aluno_id === String(alunoId) && getProcessId(choice) === CURRENT_PROCESS_ID) || null;
  },

  async getAll() {
    const client = await ensureSupabaseReady();

    if (client) {
      const { data, error } = await client
        .from("escolhas")
        .select("*")
        .eq("processo_id", CURRENT_PROCESS_ID)
        .order("submetido_em", { ascending: false });

      if (error) {
        throw new Error("Não foi possível carregar as escolhas guardadas.");
      }

      return data.map(mapChoiceFromDatabase);
    }

    return JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]")
      .filter((choice) => getProcessId(choice) === CURRENT_PROCESS_ID);
  },

  async save(choice) {
    const client = await ensureSupabaseReady();

    if (client) {
      const existingChoice = await this.getByAlunoId(choice.aluno_id);

      if (existingChoice?.estado === "bloqueada") {
        throw new Error("A submissão está bloqueada pela administração e já não pode ser alterada.");
      }

      const payload = mapChoiceToDatabase(choice, existingChoice);
      const query = existingChoice
        ? client.from("escolhas").update(payload).eq("aluno_id", String(choice.aluno_id)).eq("processo_id", CURRENT_PROCESS_ID)
        : client.from("escolhas").insert(payload);
      const { error } = await query;

      if (error) {
        throw new Error(`Não foi possível guardar a escolha: ${error.message}`);
      }

      return;
    }

    const choices = await this.getAll();
    const existingIndex = choices.findIndex((item) => item.aluno_id === choice.aluno_id);

    if (existingIndex >= 0) {
      choices[existingIndex] = choice;
    } else {
      choices.push(choice);
    }

    localStorage.setItem(STORAGE_KEY, JSON.stringify(choices));
  },

  async updateStatus(alunoId, estado) {
    const client = await ensureSupabaseReady();

    if (!client) {
      throw new Error("A alteração de estado requer ligação ao Supabase.");
    }

    const { error } = await client
      .from("escolhas")
      .update({
        estado,
        atualizado_em: new Date().toISOString()
      })
      .eq("aluno_id", String(alunoId))
      .eq("processo_id", CURRENT_PROCESS_ID);

    if (error) {
      throw new Error("Não foi possível atualizar o estado da submissão.");
    }
  },

  async clear() {
    const client = await ensureSupabaseReady();

    if (client) {
      const { error } = await client.from("escolhas").delete().eq("processo_id", CURRENT_PROCESS_ID);

      if (error) {
        throw new Error("Não foi possível limpar as escolhas na base de dados.");
      }

      return;
    }

    localStorage.removeItem(STORAGE_KEY);
  }
};

const studentStatusRepository = {
  async getAll() {
    const client = await ensureSupabaseReady();

    if (!client) {
      return new Map(JSON.parse(localStorage.getItem("escolhas12ano.alunos_estado") || "[]"));
    }

    const { data, error } = await client
      .from("alunos_estado")
      .select("aluno_id,estado,observacao,atualizado_em");

    if (error) {
      console.warn("A tabela alunos_estado ainda não está disponível.", error.message);
      return new Map();
    }

    return new Map((data || []).map((row) => [String(row.aluno_id), row.estado || "ativo"]));
  },

  async getByAlunoId(alunoId) {
    const client = await ensureSupabaseReady();

    if (!client) {
      return (await this.getAll()).get(String(alunoId)) || "ativo";
    }

    const { data, error } = await client
      .from("alunos_estado")
      .select("estado")
      .eq("aluno_id", String(alunoId))
      .maybeSingle();

    if (error) {
      console.warn("Não foi possível consultar alunos_estado.", error.message);
      return "ativo";
    }

    return data?.estado || "ativo";
  },

  async updateStatus(alunoId, estado) {
    const client = await ensureSupabaseReady();
    const payload = {
      aluno_id: String(alunoId),
      estado,
      atualizado_em: new Date().toISOString()
    };

    if (!client) {
      const statuses = await this.getAll();
      statuses.set(String(alunoId), estado);
      localStorage.setItem("escolhas12ano.alunos_estado", JSON.stringify(Array.from(statuses.entries())));
      return;
    }

    const { error } = await client
      .from("alunos_estado")
      .upsert(payload, { onConflict: "aluno_id" });

    if (error) {
      throw new Error(`Não foi possível atualizar o estado administrativo do aluno: ${error.message}`);
    }
  }
};

signInButton.addEventListener("click", async () => {
  await ensureAuthReady();

  if (!isAuthConfigured()) {
    showAuthWarning("A autenticação O365 ainda não está configurada. Preenche MSAL_CLIENT_ID em app.js.");
    return;
  }

  try {
    authWarning.classList.add("hidden");
    await msalClient.loginRedirect(loginRequest);
  } catch (error) {
    showAuthWarning(`Não foi possível iniciar sessão: ${error.message}`);
  }
});

headerSignOutButton.addEventListener("click", async () => {
  if (!msalClient || !signedInAccount) {
    return;
  }

  showHome();
  await msalClient.logoutRedirect({
    account: signedInAccount,
    postLogoutRedirectUri: HOME_URL
  });
});

homeLink.addEventListener("click", async (event) => {
  event.preventDefault();
  await showHome();
});

faqNav.addEventListener("click", (event) => {
  event.preventDefault();
  showFaqPage();
});

goHomeButton.addEventListener("click", async () => {
  await showHome();
});

editChoiceButton.addEventListener("click", () => {
  if (!currentStudent) {
    return;
  }

  showChoiceEditor();
});

cancelEditButton.addEventListener("click", () => {
  if (currentChoice) {
    renderStudentArea(currentStudent, currentChoice);
    showSummaryPage(currentChoice);
    return;
  }

  showHome();
});

faqBackButton.addEventListener("click", () => {
  if (lastStudentView === "editor") {
    showChoiceEditor();
    return;
  }

  if (currentChoice) {
    showSummaryPage(currentChoice);
    return;
  }

  showHome();
});

choicesForm.addEventListener("change", () => {
  if (isTenthGradeProcess()) {
    updateTenthGradeUi();
  } else if (isCurrentCourseCambridge()) {
    updateCambridgeUi();
  } else {
    enforcePriorityLimits();
  }

  updateValidation();
});

choicesForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const isTenthGrade = isTenthGradeProcess();
  const tenthGrade = isTenthGrade ? getTenthGradeSelection() : null;
  const isCambridge = !isTenthGrade && isCurrentCourseCambridge();
  const priorities = isTenthGrade || isCambridge ? [] : getSelectedPriorities();
  const cambridge = isCambridge ? getCambridgeSelection() : null;
  const validation = isTenthGrade
    ? validateTenthGradeSelection(tenthGrade)
    : isCambridge
    ? validateCambridgeSelection(cambridge, currentStudent.curso)
    : validatePriorities(priorities, currentStudent.curso);

  if (!validation.valid) {
    updateValidation();
    return;
  }

  const choice = {
    aluno_id: currentStudent.aluno_id,
    nome: currentStudent.nome,
    turma: currentStudent.turma,
    curso: currentStudent.curso,
    processo_id: getProcessId(currentStudent),
    ano: getProcessYear(currentStudent),
    autenticado_com: getSignedInEmail(),
    prioridades: isTenthGrade ? [] : priorities,
    cambridge,
    escolha_10: tenthGrade,
    submetido_em: currentChoice?.submetido_em || new Date().toISOString()
  };

  try {
    submitChoiceButton.disabled = true;
    await choiceRepository.save(choice);
    currentChoice = await choiceRepository.getByAlunoId(currentStudent.aluno_id);
    updateChoiceStatus(currentChoice);
    await updateCsvOutput();
    showSummaryPage(currentChoice);
  } catch (error) {
    validationMessage.textContent = error.message;
    validationMessage.className = "validation invalid";
  } finally {
    updateValidation();
  }
});

refreshResultsButton.addEventListener("click", updateAdminDashboard);
exportCsvButton.addEventListener("click", updateCsvOutput);

adminFiltersForm.addEventListener("input", () => {
  renderFilteredAdminDashboard();
});

adminFiltersForm.addEventListener("change", () => {
  updateClassFilterOptionsForSelectedCourse();
  renderFilteredAdminDashboard();
});

clearFiltersButton.addEventListener("click", () => {
  filterCourse.value = "";
  filterClass.value = "";
  filterStatus.value = "";
  filterSubmission.value = "";
  filterSearch.value = "";
  renderFilteredAdminDashboard();
});

openAdminStatsButton.addEventListener("click", () => {
  if (!isAdminUser) {
    return;
  }

  showAdminStats();
});

openAdminToolsButton.addEventListener("click", () => {
  if (!isAdminUser) {
    return;
  }

  showAdminTools();
});

backAdminFromStatsButton.addEventListener("click", () => {
  if (!isAdminUser) {
    return;
  }

  showAdminHome();
});

backAdminButton.addEventListener("click", () => {
  if (!isAdminUser) {
    return;
  }

  showAdminHome();
});

downloadCsvButton.addEventListener("click", async () => {
  const { choices } = getFilteredAdminData();
  const csv = buildChoicesCsv(choices);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = CURRENT_PROCESS.downloadName;
  link.click();
  URL.revokeObjectURL(link.href);
});

clearResultsButton.addEventListener("click", async () => {
  const confirmed = window.confirm("Queres limpar todos os resultados guardados na base de dados?");

  if (confirmed) {
    try {
      await choiceRepository.clear();
      await updateCsvOutput();
      await updateAdminDashboard();
    } catch (error) {
      csvOutput.value = error.message;
    }
  }
});

async function loadSignedInStudent(options = {}) {
  const { preferEditor = false } = options;

  if (!signedInAccount || isAdminUser) {
    return;
  }

  document.querySelector("#entrada").classList.remove("hidden");
  loginMessage.textContent = "A procurar os teus dados...";
  studentArea.classList.add("hidden");

  try {
    const student = await studentRepository.findByEmail(getSignedInEmail());

    if (!student) {
      showLoginError("Não foi encontrado um aluno associado à tua conta Microsoft 365. Contacta a secretaria ou a administração.");
      return;
    }

    if (!isTenthGradeProcess() && !getCourseRules(student.curso)) {
      showLoginError("O teu curso não está configurado nesta aplicação. Contacta a administração.");
      return;
    }

    currentStudent = student;
    const studentProcessStatus = await studentStatusRepository.getByAlunoId(student.aluno_id);

    if (studentProcessStatus === "nao_renova") {
      showLoginError("Não tens de preencher escolhas de opcionais porque a tua matrícula não está ativa para este processo. Se isto não estiver correto, contacta a secretaria.");
      return;
    }

    currentChoice = await choiceRepository.getByAlunoId(student.aluno_id);
    loginMessage.textContent = "";
    document.querySelector("#entrada").classList.add("hidden");
    renderStudentArea(student, currentChoice);

    if (preferEditor) {
      showChoiceEditor();
    } else if (currentChoice) {
      showSummaryPage(currentChoice);
    } else {
      showChoiceEditor();
    }
  } catch (error) {
    showLoginError(error.message);
  }
}

function renderStudentArea(student, existingChoice = null) {
  if (isTenthGradeProcess()) {
    renderTenthGradeStudentArea(student, existingChoice);
    return;
  }

  const rules = getCourseRules(student.curso);

  document.querySelector("#student-name").textContent = student.nome;
  document.querySelector("#student-class").textContent = student.turma;
  document.querySelector("#student-course").textContent = rules.label;
  document.querySelector("#course-rules").textContent = rules.ruleText;

  prioritiesList.innerHTML = "";
  cambridgeExtraList.innerHTML = "";
  document.querySelectorAll(".cambridge-main-card").forEach((label) => {
    label.onclick = () => {
      window.setTimeout(() => {
        updateCambridgeUi();
        updateValidation();
      }, 0);
    };
  });
  standardChoices.classList.toggle("hidden", Boolean(rules.cambridge));
  cambridgeChoices.classList.toggle("hidden", !rules.cambridge);
  tenthChoices.classList.add("hidden");

  if (rules.pendingConfiguration) {
    validationMessage.textContent = rules.ruleText;
    validationMessage.className = "validation invalid";
    submitChoiceButton.disabled = true;
    choiceStatus.textContent = "Em configuração";
    choiceStatus.className = "status-pill locked";
    confirmation.classList.add("hidden");
    studentArea.classList.remove("hidden");
    return;
  }

  if (rules.cambridge) {
    resetCambridgeForm();

    if (existingChoice?.cambridge) {
      prefillCambridgeChoice(existingChoice.cambridge);
    }

    updateCambridgeUi();
    studentArea.classList.remove("hidden");
    submitChoiceButton.disabled = true;
    updateValidation();
    return;
  }

  for (let priority = 1; priority <= 3; priority += 1) {
    prioritiesList.appendChild(createPriorityGroup(priority, rules));
  }

  if (existingChoice) {
    prefillChoice(existingChoice);
    confirmation.classList.remove("hidden");
    confirmation.textContent = existingChoice.estado === "bloqueada"
      ? "A tua submissão está bloqueada pela administração e já não pode ser alterada."
      : "Já existe uma submissão guardada. Podes rever e editar enquanto não estiver bloqueada.";
  } else {
    confirmation.classList.add("hidden");
    confirmation.textContent = "";
  }

  updateChoiceStatus(existingChoice);
  setChoicesLocked(existingChoice?.estado === "bloqueada");
  submitChoiceButton.disabled = true;
  updateValidation();
}

function renderTenthGradeStudentArea(student, existingChoice = null) {
  document.querySelector("#student-name").textContent = student.nome;
  document.querySelector("#student-class").textContent = student.turma;
  document.querySelector("#student-course").textContent = existingChoice?.escolha_10?.curso_label || "A escolher para o 10.º ano";
  document.querySelector("#course-rules").textContent = "Escolhe o curso do 10.º ano. A aplicação mostra as disciplinas associadas e valida a opção quando existir.";

  prioritiesList.innerHTML = "";
  cambridgeExtraList.innerHTML = "";
  tenthCourseOptions.innerHTML = "";
  tenthOptionList.innerHTML = "";

  standardChoices.classList.add("hidden");
  cambridgeChoices.classList.add("hidden");
  tenthChoices.classList.remove("hidden");

  Object.entries(tenthGradeCourses).forEach(([courseKey, rules]) => {
    const label = document.createElement("label");
    label.className = `subject-option ${rules.cambridge ? "optional-subject" : "required-subject"}`;

    const input = document.createElement("input");
    input.type = "radio";
    input.name = "tenth-course";
    input.value = courseKey;

    const text = document.createElement("span");
    text.textContent = rules.label;

    const note = document.createElement("small");
    note.textContent = rules.cambridge ? "Percurso Cambridge" : "Percurso nacional";
    text.appendChild(note);

    label.append(input, text);
    tenthCourseOptions.appendChild(label);
  });

  if (existingChoice?.escolha_10) {
    prefillTenthGradeChoice(existingChoice.escolha_10);
    confirmation.classList.remove("hidden");
    confirmation.textContent = existingChoice.estado === "bloqueada"
      ? "A tua submissão está bloqueada pela administração e já não pode ser alterada."
      : "Já existe uma submissão guardada. Podes rever e editar enquanto não estiver bloqueada.";
  } else {
    confirmation.classList.add("hidden");
    confirmation.textContent = "";
  }

  updateChoiceStatus(existingChoice);
  setChoicesLocked(existingChoice?.estado === "bloqueada");
  studentArea.classList.remove("hidden");
  submitChoiceButton.disabled = true;
  updateTenthGradeUi();
  updateValidation();
}

function prefillChoice(choice) {
  choice.prioridades.forEach((item) => {
    item.subjects.forEach((subject) => {
      const input = document.querySelector(`input[name="priority-${item.priority}"][value="${escapeSelectorValue(subject)}"]`);

      if (input) {
        input.checked = true;
      }
    });
  });

  enforcePriorityLimits();
}

function escapeSelectorValue(value) {
  if (window.CSS?.escape) {
    return window.CSS.escape(value);
  }

  return String(value).replaceAll('"', '\\"');
}

function updateChoiceStatus(choice) {
  if (!choice) {
    choiceStatus.textContent = "Sem submissão";
    choiceStatus.className = "status-pill";
    submitChoiceButton.textContent = "Guardar escolha";
    return;
  }

  const isLocked = choice.estado === "bloqueada";
  choiceStatus.textContent = isLocked ? "Submissão bloqueada" : "Submissão editável";
  choiceStatus.className = `status-pill ${isLocked ? "locked" : "editable"}`;
  submitChoiceButton.textContent = isLocked ? "Submissão bloqueada" : "Atualizar escolha";
}

function setChoicesLocked(isLocked) {
  if (!isLocked) {
    document.querySelectorAll('input[type="checkbox"][name^="priority-"]').forEach((input) => {
      input.disabled = false;
    });
    document.querySelectorAll('input[name="maths-october"], input[name="cambridge-main"], input[name="cambridge-extra"]').forEach((input) => {
      input.disabled = false;
    });
    document.querySelectorAll('input[name="tenth-course"], input[name="tenth-option"]').forEach((input) => {
      input.disabled = false;
    });
    enforcePriorityLimits();
    return;
  }

  document.querySelectorAll('input[type="checkbox"][name^="priority-"]').forEach((input) => {
    input.disabled = true;
  });
  document.querySelectorAll('input[name="maths-october"], input[name="cambridge-main"], input[name="cambridge-extra"]').forEach((input) => {
    input.disabled = true;
  });
  document.querySelectorAll('input[name="tenth-course"], input[name="tenth-option"]').forEach((input) => {
    input.disabled = true;
  });

  submitChoiceButton.disabled = true;
}

function isTenthGradeProcess() {
  return CURRENT_PROCESS.mode === "10-curso";
}

function getTenthGradeSelection() {
  const selectedCourseKey = document.querySelector('input[name="tenth-course"]:checked')?.value || "";
  const rules = tenthGradeCourses[selectedCourseKey] || null;
  const selectedOption = document.querySelector('input[name="tenth-option"]:checked')?.value || "";

  if (!rules) {
    return null;
  }

  const tipoCambridge = rules.cambridge
    ? "candidatura_ou_continuidade"
    : "nacional";

  return {
    curso_key: selectedCourseKey,
    curso_label: rules.label,
    cambridge: rules.cambridge,
    tipo_cambridge: tipoCambridge,
    estado_cambridge: rules.cambridge ? "entrevista" : "nao_aplicavel",
    disciplinas_automaticas: [...rules.automaticSubjects],
    disciplina_opcao: selectedOption || null,
    disciplinas: [...rules.automaticSubjects, selectedOption].filter(Boolean)
  };
}

function updateTenthGradeUi() {
  if (!isTenthGradeProcess()) {
    return;
  }

  const selection = getTenthGradeSelection();
  tenthAutomaticSubjects.classList.add("hidden");
  tenthOptionSubjects.classList.add("hidden");
  tenthCambridgeNotice.classList.add("hidden");
  tenthOptionList.innerHTML = "";

  if (!selection) {
    return;
  }

  const rules = tenthGradeCourses[selection.curso_key];
  tenthAutomaticSubjects.innerHTML = `
    <h3>Disciplinas associadas</h3>
    <p>${escapeHtml(rules.ruleText)}</p>
    <ul>${rules.automaticSubjects.map((subject) => `<li>${escapeHtml(subject)}</li>`).join("")}</ul>
  `;
  tenthAutomaticSubjects.classList.remove("hidden");
  tenthCambridgeNotice.classList.toggle("hidden", !rules.cambridge);

  if (rules.optionSubjects.length === 0) {
    return;
  }

  rules.optionSubjects.forEach((subject) => {
    const label = document.createElement("label");
    label.className = "subject-option optional-subject";

    const input = document.createElement("input");
    input.type = "radio";
    input.name = "tenth-option";
    input.value = subject;
    input.checked = subject === selection.disciplina_opcao;

    const text = document.createElement("span");
    text.textContent = subject;

    const note = document.createElement("small");
    note.textContent = "disciplina de opção";
    text.appendChild(note);

    label.append(input, text);
    tenthOptionList.appendChild(label);
  });
  tenthOptionSubjects.classList.remove("hidden");
}

function validateTenthGradeSelection(selection) {
  if (!selection) {
    return {
      valid: false,
      message: "Escolha inválida: seleciona o curso que pretendes frequentar no 10.º ano."
    };
  }

  const rules = tenthGradeCourses[selection.curso_key];

  if (!rules) {
    return {
      valid: false,
      message: "Escolha inválida: o curso selecionado não está configurado."
    };
  }

  if (rules.optionSubjects.length > 0 && !selection.disciplina_opcao) {
    return {
      valid: false,
      message: "Escolha inválida: seleciona Biologia e Geologia ou Geometria Descritiva A."
    };
  }

  if (selection.disciplina_opcao && !rules.optionSubjects.includes(selection.disciplina_opcao)) {
    return {
      valid: false,
      message: "Escolha inválida: a disciplina de opção selecionada não pertence ao curso escolhido."
    };
  }

  return {
    valid: true,
    message: rules.cambridge
      ? "Escolha válida: curso e disciplinas registados. O percurso Cambridge fica sujeito a validação pelo Colégio."
      : "Escolha válida: curso e disciplinas registados."
  };
}

function prefillTenthGradeChoice(choice) {
  const courseInput = document.querySelector(`input[name="tenth-course"][value="${escapeSelectorValue(choice.curso_key)}"]`);

  if (courseInput) {
    courseInput.checked = true;
  }

  updateTenthGradeUi();

  if (choice.disciplina_opcao) {
    const optionInput = document.querySelector(`input[name="tenth-option"][value="${escapeSelectorValue(choice.disciplina_opcao)}"]`);

    if (optionInput) {
      optionInput.checked = true;
    }
  }
}

async function showHome() {
  studentArea.classList.add("hidden");
  summaryArea.classList.add("hidden");
  faqArea.classList.add("hidden");
  document.querySelector("#entrada").classList.add("hidden");

  if (!signedInAccount) {
    document.querySelector("#autenticacao").classList.remove("hidden");
    window.history.replaceState(null, "", HOME_URL);
    window.scrollTo({ top: 0, behavior: "smooth" });
    return;
  }

  if (isAdminUser) {
    showAdminHome();
    return;
  }

  if (currentChoice) {
    showSummaryPage(currentChoice);
    return;
  }

  if (currentStudent) {
    showChoiceEditor();
    return;
  }

  await loadSignedInStudent({ preferEditor: true });
}

function showChoiceEditor() {
  if (!currentStudent) {
    return;
  }

  document.querySelector("#entrada").classList.add("hidden");
  summaryArea.classList.add("hidden");
  faqArea.classList.add("hidden");
  studentArea.classList.remove("hidden");
  lastStudentView = "editor";
  cancelEditButton.classList.toggle("hidden", !currentChoice);
  setChoicesLocked(currentChoice?.estado === "bloqueada");
  updateValidation();
  window.history.replaceState(null, "", "#opcoes");
  studentArea.scrollIntoView({ behavior: "smooth", block: "start" });
}

function showAdminHome() {
  document.querySelector("#admin-stats").classList.add("hidden");
  document.querySelector("#admin-tools").classList.add("hidden");
  document.querySelector("#admin").classList.remove("hidden");
  window.history.replaceState(null, "", "#admin");
  document.querySelector("#admin").scrollIntoView({ behavior: "smooth", block: "start" });
}

function showAdminStats() {
  document.querySelector("#admin").classList.add("hidden");
  document.querySelector("#admin-tools").classList.add("hidden");
  document.querySelector("#admin-stats").classList.remove("hidden");
  renderFilteredAdminDashboard();
  window.history.replaceState(null, "", "#admin-stats");
  document.querySelector("#admin-stats").scrollIntoView({ behavior: "smooth", block: "start" });
}

function showAdminTools() {
  document.querySelector("#admin").classList.add("hidden");
  document.querySelector("#admin-stats").classList.add("hidden");
  document.querySelector("#admin-tools").classList.remove("hidden");
  updateCsvOutput();
  window.history.replaceState(null, "", "#admin-tools");
  document.querySelector("#admin-tools").scrollIntoView({ behavior: "smooth", block: "start" });
}

function showSummaryPage(choice) {
  studentArea.classList.add("hidden");
  faqArea.classList.add("hidden");
  document.querySelector("#entrada").classList.add("hidden");
  summaryArea.classList.remove("hidden");
  lastStudentView = "summary";
  renderSummaryTable(choice);

  const isLocked = choice?.estado === "bloqueada";
  summaryMessage.textContent = isLocked
    ? "A submissão está bloqueada pela administração e não pode ser alterada."
    : "A submissão está guardada e ainda pode ser alterada.";
  editChoiceButton.classList.toggle("hidden", isLocked);
  cancelEditButton.classList.add("hidden");

  window.history.replaceState(null, "", "#resumo");
  summaryArea.scrollIntoView({ behavior: "smooth", block: "start" });
}

function showFaqPage() {
  if (!signedInAccount || isAdminUser) {
    return;
  }

  studentArea.classList.add("hidden");
  summaryArea.classList.add("hidden");
  document.querySelector("#entrada").classList.add("hidden");
  faqArea.classList.remove("hidden");
  window.history.replaceState(null, "", "#faq");
  faqArea.scrollIntoView({ behavior: "smooth", block: "start" });
}

function renderSummaryTable(choice) {
  if (!choice) {
    summaryTable.textContent = "Ainda não existe submissão.";
    return;
  }

  const table = document.createElement("table");
  table.className = "results-table";
  table.innerHTML = choice.escolha_10
    ? `
      <thead>
        <tr>
          <th>Campo</th>
          <th>Valor</th>
        </tr>
      </thead>
      <tbody></tbody>
    `
    : choice.cambridge
    ? `
      <thead>
        <tr>
          <th>Campo</th>
          <th>Valor</th>
        </tr>
      </thead>
      <tbody></tbody>
    `
    : `
      <thead>
        <tr>
          <th>Prioridade</th>
          <th>Disciplina 1</th>
          <th>Disciplina 2</th>
        </tr>
      </thead>
      <tbody></tbody>
    `;

  const tbody = table.querySelector("tbody");

  if (choice.escolha_10) {
    const rows = [
      ["Curso escolhido", choice.escolha_10.curso_label],
      ["Percurso Cambridge", choice.escolha_10.cambridge ? "Sim" : "Não"],
      ["Estado Cambridge", formatCambridgeStatus(choice.escolha_10.estado_cambridge)],
      ["Disciplinas automáticas", (choice.escolha_10.disciplinas_automaticas || []).join(" + ") || "-"],
      ["Disciplina de opção", choice.escolha_10.disciplina_opcao || "-"]
    ];

    rows.forEach(([label, value]) => {
      const row = document.createElement("tr");
      row.innerHTML = `<td>${escapeHtml(label)}</td><td>${escapeHtml(value)}</td>`;
      tbody.appendChild(row);
    });
  } else if (choice.cambridge) {
    const rows = [
      ["Exame de Maths em outubro", choice.cambridge.faz_maths_outubro ? "Sim" : "Não"],
      ["Psychology", choice.cambridge.psychology ? "Sim" : "Não"],
      ["Economics", choice.cambridge.economics ? "Sim" : "Não"],
      ["Disciplina adicional", choice.cambridge.disciplina_extra || "-"]
    ];

    rows.forEach(([label, value]) => {
      const row = document.createElement("tr");
      row.innerHTML = `<td>${escapeHtml(label)}</td><td>${escapeHtml(value)}</td>`;
      tbody.appendChild(row);
    });
  } else {
    choice.prioridades.forEach((item) => {
      const row = document.createElement("tr");
      row.innerHTML = `
        <td>Prioridade ${item.priority}</td>
        <td>${escapeHtml(item.subjects[0])}</td>
        <td>${escapeHtml(item.subjects[1])}</td>
      `;
      tbody.appendChild(row);
    });
  }

  const meta = document.createElement("dl");
  meta.className = "student-summary";
  meta.innerHTML = `
    <div>
      <dt>Aluno</dt>
      <dd>${escapeHtml(choice.nome)}</dd>
    </div>
    <div>
      <dt>Turma</dt>
      <dd>${escapeHtml(choice.turma)}</dd>
    </div>
    <div>
      <dt>Estado</dt>
      <dd>${choice.estado === "bloqueada" ? "Bloqueada" : "Editável"}</dd>
    </div>
  `;

  summaryTable.innerHTML = "";
  summaryTable.append(meta, table);
}

function isCurrentCourseCambridge() {
  return Boolean(currentStudent && getCourseRules(currentStudent.curso)?.cambridge);
}

function resetCambridgeForm() {
  document.querySelectorAll('input[name="maths-october"], input[name="cambridge-main"], input[name="cambridge-extra"]').forEach((input) => {
    input.checked = false;
    input.disabled = false;
  });

  cambridgeMainOptions.classList.add("hidden");
  cambridgeExtraOptions.classList.add("hidden");
  cambridgeAutoMessage.classList.add("hidden");
  cambridgeExtraList.innerHTML = "";
}

function getCambridgeSelection() {
  const mathsValue = document.querySelector('input[name="maths-october"]:checked')?.value || "";
  const selectedMain = Array.from(document.querySelectorAll('input[name="cambridge-main"]:checked')).map((input) => input.value);
  const extra = document.querySelector('input[name="cambridge-extra"]:checked')?.value || "";

  return {
    faz_maths_outubro: mathsValue === "yes" ? true : mathsValue === "no" ? false : null,
    psychology: mathsValue === "no" || selectedMain.includes("Psychology"),
    economics: mathsValue === "no" || selectedMain.includes("Economics"),
    disciplina_extra: extra || null
  };
}

function updateCambridgeUi() {
  const rules = getCourseRules(currentStudent.curso);
  const selection = getCambridgeSelection();
  const mathsAnswered = selection.faz_maths_outubro !== null;

  const isCtCambridge = rules.cambridgeType === "ct";
  cambridgeMainOptions.classList.toggle("hidden", !mathsAnswered || selection.faz_maths_outubro === false);
  cambridgeAutoMessage.classList.toggle("hidden", selection.faz_maths_outubro !== false);
  cambridgeAutoMessage.textContent = isCtCambridge
    ? "Ficas automaticamente inscrito em Psychology e Economics. Seleciona também uma opção d)."
    : "Ficas automaticamente inscrito em Psychology e Economics.";
  cambridgeExtraOptions.classList.add("hidden");
  cambridgeExtraList.innerHTML = "";

  if (!mathsAnswered) {
    return;
  }

  if (selection.faz_maths_outubro === false && !isCtCambridge) {
    return;
  }

  const mainCount = Number(selection.psychology) + Number(selection.economics);

  if (selection.faz_maths_outubro === true && mainCount === 0) {
    return;
  }

  const extraSubjects = getCambridgeExtraSubjects(rules, selection);

  if (extraSubjects.length === 0) {
    return;
  }

  cambridgeExtraNote.textContent = "Seleciona uma disciplina adicional.";
  extraSubjects.forEach((subject) => {
    cambridgeExtraList.appendChild(createCambridgeExtraOption(subject, rules, selection.disciplina_extra));
  });
  cambridgeExtraOptions.classList.remove("hidden");
}

function getCambridgeExtraSubjects(rules, selection) {
  if (rules.cambridgeType === "ct") {
    return ["Biologia", "Física", "Química"];
  }

  if (rules.cambridgeType === "se" && selection.psychology && selection.economics) {
    return [];
  }

  if (rules.cambridgeType === "se" && selection.psychology) {
    return ["Geografia C", "Sociologia", "Economia C"];
  }

  if (rules.cambridgeType === "se" && selection.economics) {
    return ["Geografia C", "Sociologia"];
  }

  return [];
}

function createCambridgeExtraOption(subject, rules, selectedSubject = "") {
  const label = document.createElement("label");
  label.className = "subject-option required-subject";

  const input = document.createElement("input");
  input.type = "radio";
  input.name = "cambridge-extra";
  input.value = subject;
  input.checked = subject === selectedSubject;

  const text = document.createElement("span");
  text.textContent = subject;

  const group = document.createElement("small");
  group.textContent = `${rules.requiredGroupName} - opção d)`;
  text.appendChild(group);
  label.append(input, text);
  return label;
}

function validateCambridgeSelection(selection, courseKey) {
  const rules = getCourseRules(courseKey);

  if (selection.faz_maths_outubro === null) {
    return {
      valid: false,
      message: "Escolha inválida: responde à questão sobre o exame de Maths em outubro."
    };
  }

  if (selection.faz_maths_outubro === false && rules.cambridgeType !== "ct") {
    return {
      valid: true,
      message: "Escolha válida: ficas inscrito automaticamente em Psychology e Economics."
    };
  }

  if (selection.faz_maths_outubro === false && rules.cambridgeType === "ct") {
    const extraSubjects = getCambridgeExtraSubjects(rules, selection);

    if (!selection.disciplina_extra) {
      return {
        valid: false,
        message: "Escolha inválida: seleciona uma opção d) entre Biologia, Física ou Química."
      };
    }

    if (!extraSubjects.includes(selection.disciplina_extra)) {
      return {
        valid: false,
        message: "Escolha inválida: a opção d) selecionada não é permitida."
      };
    }

    return {
      valid: true,
      message: "Escolha válida: Psychology e Economics ficam automáticos e a opção d) foi selecionada."
    };
  }

  const mainCount = Number(selection.psychology) + Number(selection.economics);

  if (mainCount === 0) {
    return {
      valid: false,
      message: "Escolha inválida: seleciona Psychology, Economics ou ambas."
    };
  }

  const extraSubjects = getCambridgeExtraSubjects(rules, selection);

  if (extraSubjects.length > 0 && !selection.disciplina_extra) {
    return {
      valid: false,
      message: "Escolha inválida: seleciona uma disciplina adicional."
    };
  }

  if (selection.disciplina_extra && !extraSubjects.includes(selection.disciplina_extra)) {
    return {
      valid: false,
      message: "Escolha inválida: a disciplina adicional selecionada não é permitida para esta combinação."
    };
  }

  return {
    valid: true,
    message: "Escolha válida para o percurso Cambridge."
  };
}

function prefillCambridgeChoice(cambridge) {
  const mathsValue = cambridge.faz_maths_outubro ? "yes" : "no";
  const mathsInput = document.querySelector(`input[name="maths-october"][value="${mathsValue}"]`);

  if (mathsInput) {
    mathsInput.checked = true;
  }

  ["Psychology", "Economics"].forEach((subject) => {
    const key = subject.toLowerCase();
    const input = document.querySelector(`input[name="cambridge-main"][value="${subject}"]`);

    if (input && cambridge[key]) {
      input.checked = true;
    }
  });

  updateCambridgeUi();

  if (cambridge.disciplina_extra) {
    const extraInput = document.querySelector(`input[name="cambridge-extra"][value="${escapeSelectorValue(cambridge.disciplina_extra)}"]`);

    if (extraInput) {
      extraInput.checked = true;
    }
  }
}

function createPriorityGroup(priority, rules) {
  const wrapper = document.createElement("section");
  wrapper.className = "priority-group";
  wrapper.setAttribute("aria-labelledby", `priority-${priority}-title`);

  const title = document.createElement("h3");
  title.id = `priority-${priority}-title`;
  title.textContent = `Prioridade ${priority}`;

  const note = document.createElement("p");
  note.className = "form-note";
  note.textContent = "Seleciona duas disciplinas para esta prioridade.";

  const subjects = document.createElement("div");
  subjects.className = "subjects-list";

  getAllowedSubjects(rules).forEach((subject) => {
    subjects.appendChild(createSubjectOption(subject, rules, priority));
  });

  wrapper.append(title, note, subjects);
  return wrapper;
}

function createSubjectOption(subject, rules, priority) {
  const label = document.createElement("label");
  label.className = "subject-option";

  if (rules.required.includes(subject)) {
    label.classList.add("required-subject");
  } else {
    label.classList.add("optional-subject");
  }

  const checkbox = document.createElement("input");
  checkbox.type = "checkbox";
  checkbox.name = `priority-${priority}`;
  checkbox.value = subject;
  checkbox.dataset.priority = String(priority);

  const text = document.createElement("span");
  text.textContent = subject;

  const group = document.createElement("small");
  group.textContent = rules.required.includes(subject) ? `${rules.requiredGroupName} - opção d)` : "opção e)";
  text.appendChild(group);

  label.append(checkbox, text);
  return label;
}

function getAllowedSubjects(rules) {
  return [...rules.required, ...rules.optional];
}

function getSelectedPriorities() {
  return [1, 2, 3].map((priority) => ({
    priority,
    subjects: Array.from(document.querySelectorAll(`input[name="priority-${priority}"]:checked`)).map((input) => input.value)
  }));
}

function enforcePriorityLimits() {
  [1, 2, 3].forEach((priority) => {
    const selectedCount = document.querySelectorAll(`input[name="priority-${priority}"]:checked`).length;
    document.querySelectorAll(`input[name="priority-${priority}"]`).forEach((input) => {
      input.disabled = !input.checked && selectedCount >= 2;
    });
  });
}

function updateValidation() {
  if (!currentStudent) {
    return;
  }

  if (currentChoice?.estado === "bloqueada") {
    validationMessage.textContent = "Submissão bloqueada pela administração.";
    validationMessage.className = "validation invalid";
    submitChoiceButton.disabled = true;
    return;
  }

  const validation = isTenthGradeProcess()
    ? validateTenthGradeSelection(getTenthGradeSelection())
    : isCurrentCourseCambridge()
    ? validateCambridgeSelection(getCambridgeSelection(), currentStudent.curso)
    : validatePriorities(getSelectedPriorities(), currentStudent.curso);
  validationMessage.textContent = validation.message;
  validationMessage.className = `validation ${validation.valid ? "valid" : "invalid"}`;
  submitChoiceButton.disabled = !validation.valid;
}

function validatePriorities(priorities, courseKey) {
  const rules = getCourseRules(courseKey);

  if (rules?.pendingConfiguration) {
    return {
      valid: false,
      message: rules.ruleText
    };
  }

  for (const item of priorities) {
    const validation = validateSelection(item.subjects, courseKey);

    if (!validation.valid) {
      return {
        valid: false,
        message: `Escolha inválida na prioridade ${item.priority}: ${validation.detail}`
      };
    }
  }

  const repeatedPriority = findRepeatedPriority(priorities);

  if (repeatedPriority) {
    return {
      valid: false,
      message: `Escolha inválida: a prioridade ${repeatedPriority.current} repete o mesmo par da prioridade ${repeatedPriority.previous}. As 3 prioridades têm de ser diferentes.`
    };
  }

  const repeatedRequiredSubject = findOverusedRequiredSubject(priorities, courseKey);

  if (repeatedRequiredSubject) {
    return {
      valid: false,
      message: `Escolha inválida: a opção d) ${repeatedRequiredSubject.subject} foi escolhida ${repeatedRequiredSubject.count} vezes. Cada opção d) só pode aparecer no máximo 2 vezes nas 3 prioridades.`
    };
  }

  return { valid: true, message: "Escolha válida: as 3 prioridades estão corretas, são diferentes e respeitam o limite das opções d)." };
}

function findRepeatedPriority(priorities) {
  const seenPairs = new Map();

  for (const item of priorities) {
    const pairKey = normalizePriorityPair(item.subjects);

    if (seenPairs.has(pairKey)) {
      return {
        previous: seenPairs.get(pairKey),
        current: item.priority
      };
    }

    seenPairs.set(pairKey, item.priority);
  }

  return null;
}

function normalizePriorityPair(subjects) {
  return [...subjects].sort((first, second) => first.localeCompare(second, "pt-PT")).join("||");
}

function normalizeCourseKey(course) {
  return courseAliases[course] || course;
}

function getCourseRules(course) {
  return courseRules[normalizeCourseKey(course)];
}

function getCourseLabel(course) {
  return getCourseRules(course)?.label || course || "Sem curso";
}

function getChoiceCourseLabel(choice) {
  return choice?.escolha_10?.curso_label || getCourseLabel(choice?.curso);
}

function formatCambridgeStatus(status) {
  const labels = {
    entrevista: "Sujeito a entrevista/verificação",
    aplicavel: "Aplicável",
    nao_aplicavel: "Não aplicável"
  };

  return labels[status] || status || "Não aplicável";
}

function findOverusedRequiredSubject(priorities, courseKey) {
  const rules = getCourseRules(courseKey);
  const counts = new Map();

  priorities.flatMap((item) => item.subjects).forEach((subject) => {
    if (!rules.required.includes(subject)) {
      return;
    }

    counts.set(subject, (counts.get(subject) || 0) + 1);
  });

  for (const [subject, count] of counts) {
    if (count > 2) {
      return { subject, count };
    }
  }

  return null;
}

function validateSelection(selectedSubjects, courseKey) {
  const normalizedCourseKey = normalizeCourseKey(courseKey);
  const rules = getCourseRules(courseKey);

  if (selectedSubjects.length !== 2) {
    return {
      valid: false,
      detail: `seleciona exatamente duas disciplinas. Tens ${selectedSubjects.length} selecionada(s).`
    };
  }

  const hasRequiredSubject = selectedSubjects.some((subject) => rules.required.includes(subject));

  if (!hasRequiredSubject) {
    return {
      valid: false,
      detail: `pelo menos uma disciplina tem de pertencer às ${rules.requiredGroupName}.`
    };
  }

  return { valid: true };
}

function parseCsv(text) {
  const lines = text.trim().split(/\r?\n/);
  const headers = splitCsvLine(lines.shift());

  return lines.map((line) => {
    const values = splitCsvLine(line);
    return headers.reduce((student, header, index) => {
      student[header] = values[index] || "";
      return student;
    }, {});
  });
}

// Parser CSV pequeno para suportar nomes com vírgulas se forem colocados entre aspas.
function splitCsvLine(line) {
  const values = [];
  let current = "";
  let insideQuotes = false;

  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];
    const nextChar = line[index + 1];

    if (char === '"' && nextChar === '"') {
      current += '"';
      index += 1;
    } else if (char === '"') {
      insideQuotes = !insideQuotes;
    } else if (char === "," && !insideQuotes) {
      values.push(current.trim());
      current = "";
    } else {
      current += char;
    }
  }

  values.push(current.trim());
  return values;
}

function showLoginError(message) {
  loginMessage.textContent = message;
  studentArea.classList.add("hidden");
}

async function ensureSupabaseReady() {
  if (!supabaseInitPromise) {
    supabaseInitPromise = initSupabase();
  }

  return supabaseInitPromise;
}

async function initSupabase() {
  if (!SUPABASE_URL || !SUPABASE_ANON_KEY) {
    return null;
  }

  const supabaseModule = await loadSupabaseModule();

  if (!supabaseModule) {
    return null;
  }

  supabaseClient = supabaseModule.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
  return supabaseClient;
}

async function loadSupabaseModule() {
  for (const url of SUPABASE_MODULE_URLS) {
    try {
      return await import(url);
    } catch (error) {
      console.warn("Falha ao carregar Supabase:", url, error);
    }
  }

  return null;
}

function normalizeStudent(student) {
  return {
    aluno_id: String(student.aluno_id || ""),
    nome: student.nome || "",
    turma: student.turma || "",
    curso: student.curso || "",
    email: student.email || "",
    processo_id: getProcessId(student),
    ano: getProcessYear(student)
  };
}

function getProcessId(record) {
  return record?.processo_id || CURRENT_PROCESS_ID;
}

function getProcessYear(record) {
  return record?.ano || CURRENT_PROCESS_YEAR;
}

function mapChoiceToDatabase(choice, existingChoice = null) {
  const prioritySubjects = getPrioritySubjectsForExport(choice);
  const now = new Date().toISOString();
  const cambridge = choice.cambridge || {};
  const tenthGrade = choice.escolha_10 || null;

  const payload = {
    aluno_id: String(choice.aluno_id),
    nome: choice.nome,
    turma: choice.turma,
    curso: choice.curso,
    processo_id: getProcessId(choice),
    ano: getProcessYear(choice),
    email_autenticado: choice.autenticado_com,
    prioridade_1_disciplina_1: prioritySubjects[0] || "",
    prioridade_1_disciplina_2: prioritySubjects[1] || "",
    prioridade_2_disciplina_1: prioritySubjects[2] || "",
    prioridade_2_disciplina_2: prioritySubjects[3] || "",
    prioridade_3_disciplina_1: prioritySubjects[4] || "",
    prioridade_3_disciplina_2: prioritySubjects[5] || "",
    faz_maths_outubro: cambridge.faz_maths_outubro ?? null,
    cambridge_psychology: Boolean(cambridge.psychology),
    cambridge_economics: Boolean(cambridge.economics),
    cambridge_disciplina_extra: cambridge.disciplina_extra || null,
    submetido_em: existingChoice?.submetido_em || choice.submetido_em,
    atualizado_em: existingChoice ? now : null,
    estado: existingChoice?.estado || "submetida"
  };

  if (tenthGrade) {
    payload.curso_10 = tenthGrade.curso_label;
    payload.escolha_10_curso_key = tenthGrade.curso_key;
    payload.cambridge_10 = Boolean(tenthGrade.cambridge);
    payload.tipo_cambridge = tenthGrade.tipo_cambridge || null;
    payload.estado_cambridge = tenthGrade.estado_cambridge || null;
    payload.disciplinas_automaticas_10 = (tenthGrade.disciplinas_automaticas || []).join(" | ");
    payload.disciplina_opcao_10 = tenthGrade.disciplina_opcao || null;
  }

  return payload;
}

function mapChoiceFromDatabase(row) {
  const hasCambridgeData =
    (row.faz_maths_outubro !== null && row.faz_maths_outubro !== undefined) ||
    row.cambridge_psychology ||
    row.cambridge_economics ||
    row.cambridge_disciplina_extra;
  const hasTenthGradeData = Boolean(row.escolha_10_curso_key || row.curso_10 || row.disciplina_opcao_10);

  return {
    aluno_id: row.aluno_id,
    nome: row.nome,
    turma: row.turma,
    curso: row.curso,
    processo_id: getProcessId(row),
    ano: getProcessYear(row),
    autenticado_com: row.email_autenticado,
    prioridades: [
      {
        priority: 1,
        subjects: [row.prioridade_1_disciplina_1, row.prioridade_1_disciplina_2].filter(Boolean)
      },
      {
        priority: 2,
        subjects: [row.prioridade_2_disciplina_1, row.prioridade_2_disciplina_2].filter(Boolean)
      },
      {
        priority: 3,
        subjects: [row.prioridade_3_disciplina_1, row.prioridade_3_disciplina_2].filter(Boolean)
      }
    ],
    cambridge: hasCambridgeData
      ? {
          faz_maths_outubro: row.faz_maths_outubro,
          psychology: Boolean(row.cambridge_psychology),
          economics: Boolean(row.cambridge_economics),
          disciplina_extra: row.cambridge_disciplina_extra || null
        }
      : null,
    escolha_10: hasTenthGradeData
      ? {
          curso_key: row.escolha_10_curso_key || "",
          curso_label: row.curso_10 || tenthGradeCourses[row.escolha_10_curso_key]?.label || "",
          cambridge: Boolean(row.cambridge_10),
          tipo_cambridge: row.tipo_cambridge || null,
          estado_cambridge: row.estado_cambridge || null,
          disciplinas_automaticas: String(row.disciplinas_automaticas_10 || "").split("|").map((subject) => subject.trim()).filter(Boolean),
          disciplina_opcao: row.disciplina_opcao_10 || null,
          disciplinas: [
            ...String(row.disciplinas_automaticas_10 || "").split("|").map((subject) => subject.trim()).filter(Boolean),
            row.disciplina_opcao_10
          ].filter(Boolean)
        }
      : null,
    submetido_em: row.submetido_em,
    atualizado_em: row.atualizado_em,
    estado: row.estado || "submetida"
  };
}

async function initAuth() {
  signInButton.disabled = true;

  if (!isClientIdConfigured()) {
    showAuthWarning("Autenticação O365 por configurar: regista a app no Microsoft Entra e substitui MSAL_CLIENT_ID em app.js.");
    updateAuthUi();
    return;
  }

  const msalLoaded = await loadMsalLibrary();

  if (!msalLoaded) {
    showAuthWarning("Não foi possível carregar a biblioteca de autenticação Microsoft. Confirma a ligação à internet ou tenta novamente.");
    updateAuthUi();
    return;
  }

  const msalConfig = {
    auth: {
      clientId: MSAL_CLIENT_ID,
      authority: `https://login.microsoftonline.com/${MSAL_TENANT}`,
      redirectUri: MSAL_REDIRECT_URI
    },
    cache: {
      cacheLocation: "sessionStorage",
      storeAuthStateInCookie: false
    }
  };

  msalClient = new msal.PublicClientApplication(msalConfig);

  try {
    const response = await msalClient.handleRedirectPromise();

    if (response && response.account) {
      signedInAccount = response.account;
    } else {
      const accounts = msalClient.getAllAccounts();
      signedInAccount = accounts[0] || null;
    }

    if (signedInAccount) {
      msalClient.setActiveAccount(signedInAccount);
      isAdminUser = await checkAdminAccess(getSignedInEmail());
    }
  } catch (error) {
    showAuthWarning(`Erro ao concluir autenticação: ${error.message}`);
  }

  updateAuthUi();
}

async function ensureAuthReady() {
  if (!authInitPromise) {
    authInitPromise = initAuth();
  }

  await authInitPromise;
}

async function loadMsalLibrary() {
  if (window.msal) {
    return true;
  }

  for (const url of MSAL_SCRIPT_URLS) {
    const loaded = await loadScriptWithTimeout(url, 7000);

    if (loaded && window.msal) {
      return true;
    }
  }

  return false;
}

function loadScriptWithTimeout(url, timeoutMs) {
  return new Promise((resolve) => {
    const script = document.createElement("script");
    const timer = window.setTimeout(() => {
      script.remove();
      resolve(false);
    }, timeoutMs);

    script.src = url;
    script.async = true;
    script.onload = () => {
      window.clearTimeout(timer);
      resolve(true);
    };
    script.onerror = () => {
      window.clearTimeout(timer);
      script.remove();
      resolve(false);
    };

    document.head.appendChild(script);
  });
}

function isClientIdConfigured() {
  return MSAL_CLIENT_ID && !MSAL_CLIENT_ID.includes("COLOCA_AQUI");
}

function isAuthConfigured() {
  return window.msal && isClientIdConfigured() && Boolean(msalClient);
}

async function checkAdminAccess(email) {
  const client = await ensureSupabaseReady();

  if (!client || !email) {
    return false;
  }

  const { data, error } = await client
    .from("administradores")
    .select("email")
    .eq("email", email.toLowerCase())
    .eq("ativo", true)
    .maybeSingle();

  if (error) {
    console.warn("Não foi possível validar administrador:", error);
    return false;
  }

  return Boolean(data);
}

function updateAuthUi() {
  const isSignedIn = Boolean(signedInAccount);
  const isAdmin = isSignedIn && isAdminUser;

  document.querySelector("#autenticacao").classList.toggle("hidden", isSignedIn);
  document.querySelectorAll(".requires-auth").forEach((element) => {
    element.classList.toggle("hidden", !isSignedIn || isAdmin);
  });

  document.querySelectorAll(".admin-only").forEach((element) => {
    element.classList.toggle("hidden", !isAdmin);
  });

  adminNav.classList.toggle("hidden", !isAdmin);
  faqNav.classList.toggle("hidden", !isSignedIn || isAdmin);
  headerSession.classList.toggle("hidden", !isSignedIn);
  signInButton.disabled = false;

  if (isSignedIn) {
    headerAuthEmail.textContent = getSignedInDisplayName();
    loadProfilePhoto();

    if (isAdmin) {
      studentArea.classList.add("hidden");
      summaryArea.classList.add("hidden");
      showAdminHome();
      updateCsvOutput();
      updateAdminDashboard();
    } else {
      loadSignedInStudent();
    }
  } else {
    isAdminUser = false;
    headerAuthEmail.textContent = "";
    profilePhoto.classList.add("hidden");
    profilePhoto.removeAttribute("src");
    studentArea.classList.add("hidden");
    summaryArea.classList.add("hidden");
    faqArea.classList.add("hidden");
    document.querySelector("#admin").classList.add("hidden");
    document.querySelector("#admin-tools").classList.add("hidden");
    document.querySelector("#entrada").classList.add("hidden");
    window.history.replaceState(null, "", HOME_URL);
  }
}

function showAuthWarning(message) {
  authWarning.textContent = message;
  authWarning.classList.remove("hidden");
}

function getSignedInEmail() {
  if (!signedInAccount) {
    return "";
  }

  return (
    signedInAccount.username ||
    signedInAccount.idTokenClaims?.preferred_username ||
    signedInAccount.idTokenClaims?.email ||
    ""
  ).toLowerCase();
}

function getSignedInDisplayName() {
  if (!signedInAccount) {
    return "";
  }

  return (
    signedInAccount.name ||
    signedInAccount.idTokenClaims?.name ||
    getSignedInEmail()
  );
}

async function loadProfilePhoto() {
  if (!msalClient || !signedInAccount) {
    return;
  }

  try {
    const tokenResponse = await msalClient.acquireTokenSilent({
      ...loginRequest,
      account: signedInAccount
    });
    const response = await fetch("https://graph.microsoft.com/v1.0/me/photo/$value", {
      headers: {
        Authorization: `Bearer ${tokenResponse.accessToken}`
      }
    });

    if (!response.ok) {
      showInitialsAvatar();
      return;
    }

    const blob = await response.blob();
    profilePhoto.src = URL.createObjectURL(blob);
    profilePhoto.alt = `Foto de ${getSignedInDisplayName()}`;
    profilePhoto.classList.remove("hidden");
  } catch (error) {
    showInitialsAvatar();
  }
}

function showInitialsAvatar() {
  const name = getSignedInDisplayName();
  const initials = name
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => part[0].toUpperCase())
    .join("") || "?";

  const svg = `
    <svg xmlns="http://www.w3.org/2000/svg" width="96" height="96" viewBox="0 0 96 96">
      <rect width="96" height="96" rx="48" fill="#2f6f73"/>
      <text x="48" y="57" text-anchor="middle" font-family="Arial, Helvetica, sans-serif" font-size="34" font-weight="700" fill="#ffffff">${initials}</text>
    </svg>
  `;

  profilePhoto.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
  profilePhoto.alt = `Perfil de ${name}`;
  profilePhoto.classList.remove("hidden");
}

function studentMatchesSignedInAccount(student) {
  if (!student.email) {
    return true;
  }

  return student.email.toLowerCase() === getSignedInEmail();
}

async function updateCsvOutput() {
  try {
    const choices = adminChoicesCache.length ? getFilteredAdminData().choices : await choiceRepository.getAll();
    const csv = buildChoicesCsv(choices);
    csvOutput.value = csv || "Sem resultados submetidos.";
  } catch (error) {
    csvOutput.value = error.message;
  }
}

async function updateAdminDashboard() {
  if (!adminDashboard) {
    return;
  }

  adminDashboard.textContent = "A carregar indicadores...";

  try {
    const [studentsList, choices, studentStatuses] = await Promise.all([
      studentRepository.getAll(),
      choiceRepository.getAll(),
      studentStatusRepository.getAll()
    ]);
    adminStudentsCache = studentsList;
    adminChoicesCache = choices;
    adminStudentStatusesCache = studentStatuses;
    updateAdminFilterOptions(studentsList);
    renderFilteredAdminDashboard();
  } catch (error) {
    adminDashboard.textContent = error.message;
    adminResults.textContent = error.message;
  }
}

function renderAdminDashboard(studentsList, choices) {
  const submittedIds = new Set(choices.map((choice) => String(choice.aluno_id)));
  const notRenewingStudents = studentsList.filter((student) => isStudentNotRenewing(student.aluno_id));
  const pendingStudents = studentsList.filter((student) => !submittedIds.has(String(student.aluno_id)) && !isStudentNotRenewing(student.aluno_id));
  const lockedCount = choices.filter((choice) => choice.estado === "bloqueada").length;
  const editableCount = choices.length - lockedCount;

  adminDashboard.innerHTML = "";
  const dashboardRow = document.createElement("div");
  dashboardRow.className = "metrics-with-action";
  dashboardRow.appendChild(
    createMetricGrid([
      ["Alunos", studentsList.length],
      ["Submissões", choices.length],
      ["Por preencher", pendingStudents.length],
      ["Não renovam", notRenewingStudents.length],
      ["Editáveis", editableCount],
      ["Bloqueadas", lockedCount]
    ])
  );
  dashboardRow.appendChild(refreshResultsButton);
  adminDashboard.appendChild(dashboardRow);
}

function renderAdminStats(studentsList, choices) {
  if (!adminStatsDashboard) {
    return;
  }

  const submittedIds = new Set(choices.map((choice) => String(choice.aluno_id)));
  const pendingStudents = studentsList.filter((student) => !submittedIds.has(String(student.aluno_id)) && !isStudentNotRenewing(student.aluno_id));
  const notRenewingStudents = studentsList.filter((student) => isStudentNotRenewing(student.aluno_id));
  const byCourse = buildCourseStats(studentsList, choices);
  const subjectStats = buildSubjectStats(choices);
  const subjectByCourse = buildSubjectStatsByCourse(choices);
  const priorityStats = buildPrioritySubjectStats(choices);
  const pairStats = buildPairStats(choices);
  const pairStatsByPriority = buildPairStatsByPriority(choices);
  const pairStatsByCourse = buildPairStatsByCourse(choices);
  const tabs = [
    ["submissoes", "Submissões por curso", createCourseStatsTable(byCourse)],
    ["pendentes", "Alunos por preencher", createPendingStudentsTable(pendingStudents)],
    ["nao-renovam", "Não renovam", createNotRenewingStudentsTable(notRenewingStudents)],
    ["disciplinas", "Disciplinas mais escolhidas", createSubjectStatsTable(subjectStats)],
    ["disciplinas-curso", "Disciplinas mais escolhidas por curso", createSubjectStatsByCourseTable(subjectByCourse)],
    ["prioridades", "Prioridades", createPrioritySubjectStatsTable(priorityStats)],
    ["pares", "Pares escolhidos", createPairStatsTable(pairStats)],
    ["pares-prioridade", "Pares por prioridade", createPairStatsByPriorityTable(pairStatsByPriority)],
    ["pares-curso", "Pares por curso", createPairStatsByCourseTable(pairStatsByCourse)]
  ];

  adminStatsDashboard.innerHTML = "";
  adminStatsDashboard.appendChild(createStatsTabs(tabs));
}

function createStatsTabs(tabs) {
  const wrapper = document.createElement("div");
  wrapper.className = "stats-tabs";

  const tabList = document.createElement("div");
  tabList.className = "stats-tab-list";
  tabList.setAttribute("role", "tablist");
  tabList.setAttribute("aria-label", "Estatísticas disponíveis");

  const panels = document.createElement("div");
  panels.className = "stats-tab-panels";

  tabs.forEach(([id, title, content], index) => {
    const isActive = index === 0;
    const tabId = `stats-tab-${id}`;
    const panelId = `stats-panel-${id}`;

    const button = document.createElement("button");
    button.type = "button";
    button.className = `stats-tab${isActive ? " active" : ""}`;
    button.id = tabId;
    button.textContent = title;
    button.setAttribute("role", "tab");
    button.setAttribute("aria-selected", String(isActive));
    button.setAttribute("aria-controls", panelId);

    const panel = createAdminSection(title, content);
    panel.classList.add("stats-tab-panel");
    panel.id = panelId;
    panel.setAttribute("role", "tabpanel");
    panel.setAttribute("aria-labelledby", tabId);
    panel.classList.toggle("hidden", !isActive);

    button.addEventListener("click", () => {
      tabList.querySelectorAll(".stats-tab").forEach((tab) => {
        tab.classList.toggle("active", tab === button);
        tab.setAttribute("aria-selected", String(tab === button));
      });

      panels.querySelectorAll(".stats-tab-panel").forEach((item) => {
        item.classList.toggle("hidden", item !== panel);
      });
    });

    tabList.appendChild(button);
    panels.appendChild(panel);
  });

  wrapper.append(tabList, panels);
  return wrapper;
}

function renderFilteredAdminDashboard() {
  const filtered = getFilteredAdminData();
  renderAdminDashboard(filtered.students, filtered.choices);
  renderAdminStats(filtered.students, filtered.choices);
  updateAdminResults(filtered.students, filtered.choices);
  csvOutput.value = buildChoicesCsv(filtered.choices) || "Sem resultados submetidos.";
}

function getFilteredAdminData() {
  const filters = getAdminFilters();
  const choicesByAluno = new Map(adminChoicesCache.map((choice) => [String(choice.aluno_id), choice]));

  const students = adminStudentsCache.filter((student) => studentMatchesAdminFilters(student, choicesByAluno.get(String(student.aluno_id)), filters));
  const visibleStudentIds = new Set(students.map((student) => String(student.aluno_id)));
  const choices = adminChoicesCache.filter((choice) => visibleStudentIds.has(String(choice.aluno_id)) && !isStudentNotRenewing(choice.aluno_id));

  return { students, choices };
}

function getAdminFilters() {
  return {
    course: filterCourse.value,
    className: filterClass.value,
    status: filterStatus.value,
    submission: filterSubmission.value,
    search: filterSearch.value.trim().toLowerCase()
  };
}

function studentMatchesAdminFilters(student, choice, filters) {
  const submitted = Boolean(choice);
  const notRenewing = isStudentNotRenewing(student.aluno_id);

  if (filters.course && (choice ? getChoiceCourseLabel(choice) : getCourseLabel(student.curso)) !== filters.course) {
    return false;
  }

  if (filters.className && student.turma !== filters.className) {
    return false;
  }

  if (filters.status && (!choice || notRenewing || choice.estado !== filters.status)) {
    return false;
  }

  if (filters.submission === "submitted" && (!submitted || notRenewing)) {
    return false;
  }

  if (filters.submission === "pending" && (submitted || notRenewing)) {
    return false;
  }

  if (filters.submission === "not-renewing" && !notRenewing) {
    return false;
  }

  if (filters.search) {
    const haystack = [
      student.aluno_id,
      student.nome,
      student.turma,
      student.curso,
      student.email,
      choice?.email_autenticado,
      choice?.autenticado_com
    ]
      .filter(Boolean)
      .join(" ")
      .toLowerCase();

    return haystack.includes(filters.search);
  }

  return true;
}

function isStudentNotRenewing(alunoId) {
  return adminStudentStatusesCache.get(String(alunoId)) === "nao_renova";
}

function updateAdminFilterOptions(studentsList) {
  const currentCourse = filterCourse.value;
  const currentClass = filterClass.value;
  const studentCourses = studentsList.map((student) => getCourseLabel(student.curso));
  const choiceCourses = adminChoicesCache.map((choice) => getChoiceCourseLabel(choice));
  fillSelectOptions(filterCourse, uniqueSorted([...studentCourses, ...choiceCourses]));
  filterCourse.value = Array.from(filterCourse.options).some((option) => option.value === currentCourse) ? currentCourse : "";
  updateClassFilterOptionsForSelectedCourse(currentClass);
}

function updateClassFilterOptionsForSelectedCourse(preferredClass = filterClass.value) {
  const selectedCourse = filterCourse.value;
  const choicesByAluno = new Map(adminChoicesCache.map((choice) => [String(choice.aluno_id), choice]));
  const availableStudents = selectedCourse
    ? adminStudentsCache.filter((student) => {
        const choice = choicesByAluno.get(String(student.aluno_id));
        return (choice ? getChoiceCourseLabel(choice) : getCourseLabel(student.curso)) === selectedCourse;
      })
    : adminStudentsCache;
  fillSelectOptions(filterClass, uniqueSorted(availableStudents.map((student) => student.turma)));
  const currentClass = preferredClass;
  filterClass.value = Array.from(filterClass.options).some((option) => option.value === currentClass) ? currentClass : "";
}

function fillSelectOptions(select, values) {
  const firstOption = select.querySelector("option");
  select.innerHTML = "";
  select.appendChild(firstOption);

  values.forEach((value) => {
    if (!value) {
      return;
    }

    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });
}

function uniqueSorted(values) {
  return Array.from(new Set(values.filter(Boolean))).sort((first, second) => first.localeCompare(second, "pt-PT"));
}

function createMetricGrid(items) {
  const grid = document.createElement("div");
  grid.className = "metric-grid";

  items.forEach(([label, value]) => {
    const card = document.createElement("article");
    card.className = "metric-card";
    card.innerHTML = `
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
    `;
    grid.appendChild(card);
  });

  return grid;
}

function createAdminSection(title, content) {
  const section = document.createElement("section");
  section.className = "admin-section";

  const heading = document.createElement("h3");
  heading.textContent = title;

  section.append(heading, content);
  return section;
}

function buildCourseStats(studentsList, choices) {
  const submittedByCourse = new Map();
  const totalsByCourse = new Map();
  const notRenewingByCourse = new Map();

  studentsList.forEach((student) => {
    const label = getCourseLabel(student.curso);
    totalsByCourse.set(label, (totalsByCourse.get(label) || 0) + 1);

    if (isStudentNotRenewing(student.aluno_id)) {
      notRenewingByCourse.set(label, (notRenewingByCourse.get(label) || 0) + 1);
    }
  });

  choices.forEach((choice) => {
    if (isStudentNotRenewing(choice.aluno_id)) {
      return;
    }

    const label = getChoiceCourseLabel(choice);
    submittedByCourse.set(label, (submittedByCourse.get(label) || 0) + 1);
  });

  return Array.from(totalsByCourse.entries())
    .map(([course, total]) => {
      const submitted = submittedByCourse.get(course) || 0;
      const notRenewing = notRenewingByCourse.get(course) || 0;
      return {
        course,
        total,
        submitted,
        pending: total - submitted - notRenewing,
        notRenewing
      };
    })
    .sort((first, second) => first.course.localeCompare(second.course, "pt-PT"));
}

function buildSubjectStats(choices) {
  const counts = new Map();

  choices.forEach((choice) => {
    const uniqueSubjects = new Set(getPrioritySubjectsForExport(choice).filter(Boolean));

    uniqueSubjects.forEach((subject) => {
      if (!subject) {
        return;
      }

      counts.set(subject, (counts.get(subject) || 0) + 1);
    });
  });

  return Array.from(counts.entries())
    .map(([subject, count]) => ({ subject, count }))
    .sort((first, second) => second.count - first.count || first.subject.localeCompare(second.subject, "pt-PT"));
}

function buildSubjectStatsByCourse(choices) {
  const counts = new Map();

  choices.forEach((choice) => {
    const course = getChoiceCourseLabel(choice);
    const uniqueSubjects = new Set(getPrioritySubjectsForExport(choice).filter(Boolean));

    uniqueSubjects.forEach((subject) => {
      if (!subject) {
        return;
      }

      const key = `${course}||${subject}`;
      counts.set(key, {
        course,
        subject,
        count: (counts.get(key)?.count || 0) + 1
      });
    });
  });

  return Array.from(counts.values())
    .sort((first, second) =>
      first.course.localeCompare(second.course, "pt-PT") ||
      second.count - first.count ||
      first.subject.localeCompare(second.subject, "pt-PT")
    );
}

function getNormalPriorityItems(choice) {
  if (!Array.isArray(choice.prioridades)) {
    return [];
  }

  return choice.prioridades
    .filter((item) => Array.isArray(item.subjects) && item.subjects.filter(Boolean).length === 2)
    .map((item) => ({
      priority: Number(item.priority) || 0,
      subjects: item.subjects.filter(Boolean)
    }));
}

function normalizePairLabel(subjects) {
  return [...subjects]
    .sort((first, second) => first.localeCompare(second, "pt-PT"))
    .join(" + ");
}

function buildPrioritySubjectStats(choices) {
  const counts = new Map();

  choices.forEach((choice) => {
    getNormalPriorityItems(choice).forEach((item) => {
      item.subjects.forEach((subject) => {
        const key = `${item.priority}||${subject}`;
        counts.set(key, {
          priority: item.priority,
          subject,
          count: (counts.get(key)?.count || 0) + 1
        });
      });
    });
  });

  return Array.from(counts.values()).sort((first, second) =>
    first.priority - second.priority ||
    second.count - first.count ||
    first.subject.localeCompare(second.subject, "pt-PT")
  );
}

function buildPairStats(choices) {
  const counts = new Map();

  choices.forEach((choice) => {
    getNormalPriorityItems(choice).forEach((item) => {
      const pair = normalizePairLabel(item.subjects);
      counts.set(pair, (counts.get(pair) || 0) + 1);
    });
  });

  return Array.from(counts.entries())
    .map(([pair, count]) => ({ pair, count }))
    .sort((first, second) => second.count - first.count || first.pair.localeCompare(second.pair, "pt-PT"));
}

function buildPairStatsByPriority(choices) {
  const counts = new Map();

  choices.forEach((choice) => {
    getNormalPriorityItems(choice).forEach((item) => {
      const pair = normalizePairLabel(item.subjects);
      const key = `${item.priority}||${pair}`;
      counts.set(key, {
        priority: item.priority,
        pair,
        count: (counts.get(key)?.count || 0) + 1
      });
    });
  });

  return Array.from(counts.values()).sort((first, second) =>
    first.priority - second.priority ||
    second.count - first.count ||
    first.pair.localeCompare(second.pair, "pt-PT")
  );
}

function buildPairStatsByCourse(choices) {
  const counts = new Map();

  choices.forEach((choice) => {
    const course = getChoiceCourseLabel(choice);

    getNormalPriorityItems(choice).forEach((item) => {
      const pair = normalizePairLabel(item.subjects);
      const key = `${course}||${pair}`;
      counts.set(key, {
        course,
        pair,
        count: (counts.get(key)?.count || 0) + 1
      });
    });
  });

  return Array.from(counts.values()).sort((first, second) =>
    first.course.localeCompare(second.course, "pt-PT") ||
    second.count - first.count ||
    first.pair.localeCompare(second.pair, "pt-PT")
  );
}

function createCourseStatsTable(rows) {
  if (rows.length === 0) {
    return createEmptyMessage("Sem alunos registados.");
  }

  return createTable(
    ["Curso", "Alunos", "Preencheram", "Por preencher", "Não renovam"],
    rows.map((row) => [row.course, row.total, row.submitted, row.pending, row.notRenewing])
  );
}

function createPendingStudentsTable(studentsList) {
  if (studentsList.length === 0) {
    return createEmptyMessage("Todos os alunos registados já preencheram.");
  }

  return createTable(
    ["Aluno", "Turma", "Curso", "Email"],
    studentsList
      .sort((first, second) => first.turma.localeCompare(second.turma, "pt-PT") || first.nome.localeCompare(second.nome, "pt-PT"))
      .map((student) => [student.nome, student.turma, getCourseLabel(student.curso), student.email || "-"])
  );
}

function createNotRenewingStudentsTable(studentsList) {
  if (studentsList.length === 0) {
    return createEmptyMessage("Ainda não existem alunos marcados como não renovam.");
  }

  return createTable(
    ["Aluno", "Turma", "Curso", "Email"],
    studentsList
      .sort((first, second) => first.turma.localeCompare(second.turma, "pt-PT") || first.nome.localeCompare(second.nome, "pt-PT"))
      .map((student) => [student.nome, student.turma, getCourseLabel(student.curso), student.email || "-"])
  );
}

function createSubjectStatsTable(rows) {
  if (rows.length === 0) {
    return createEmptyMessage("Ainda não existem disciplinas escolhidas.");
  }

  return createTable(
    ["Disciplina", "Alunos"],
    rows.map((row) => [row.subject, row.count])
  );
}

function createSubjectStatsByCourseTable(rows) {
  if (rows.length === 0) {
    return createEmptyMessage("Ainda não existem disciplinas escolhidas por curso.");
  }

  return createTable(
    ["Curso", "Disciplina", "Alunos"],
    rows.map((row) => [row.course, row.subject, row.count])
  );
}

function createPrioritySubjectStatsTable(rows) {
  if (rows.length === 0) {
    return createEmptyMessage("Ainda não existem prioridades registadas nos cursos com prioridades.");
  }

  return createTable(
    ["Prioridade", "Disciplina", "Ocorrências"],
    rows.map((row) => [`Prioridade ${row.priority}`, row.subject, row.count])
  );
}

function createPairStatsTable(rows) {
  if (rows.length === 0) {
    return createEmptyMessage("Ainda não existem pares registados nos cursos com prioridades.");
  }

  return createTable(
    ["Par de disciplinas", "Ocorrências"],
    rows.map((row) => [row.pair, row.count])
  );
}

function createPairStatsByPriorityTable(rows) {
  if (rows.length === 0) {
    return createEmptyMessage("Ainda não existem pares por prioridade registados.");
  }

  return createTable(
    ["Prioridade", "Par de disciplinas", "Ocorrências"],
    rows.map((row) => [`Prioridade ${row.priority}`, row.pair, row.count])
  );
}

function createPairStatsByCourseTable(rows) {
  if (rows.length === 0) {
    return createEmptyMessage("Ainda não existem pares por curso registados.");
  }

  return createTable(
    ["Curso", "Par de disciplinas", "Ocorrências"],
    rows.map((row) => [row.course, row.pair, row.count])
  );
}

function createTable(headers, rows) {
  const wrapper = document.createElement("div");
  wrapper.className = "admin-table-block";

  const tools = document.createElement("div");
  tools.className = "table-tools";

  const search = document.createElement("input");
  search.type = "search";
  search.placeholder = "Filtrar nesta tabela";
  search.setAttribute("aria-label", "Filtrar nesta tabela");
  tools.appendChild(search);

  const tableWrap = document.createElement("div");
  tableWrap.className = "admin-table-wrap";
  const table = document.createElement("table");
  table.className = "results-table";
  const thead = document.createElement("thead");
  const tbody = document.createElement("tbody");
  let currentRows = [...rows];
  let sortIndex = null;
  let sortDirection = "asc";

  const headerRow = document.createElement("tr");
  headers.forEach((header, index) => {
    const th = document.createElement("th");
    const button = document.createElement("button");
    button.type = "button";
    button.className = "table-sort";
    button.textContent = header;
    button.setAttribute("aria-label", `Ordenar por ${header}`);
    button.addEventListener("click", () => {
      if (sortIndex === index) {
        sortDirection = sortDirection === "asc" ? "desc" : "asc";
      } else {
        sortIndex = index;
        sortDirection = "asc";
      }

      renderTableRows();
    });
    th.appendChild(button);
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  search.addEventListener("input", renderTableRows);

  function renderTableRows() {
    const query = search.value.trim().toLowerCase();
    currentRows = rows.filter((row) => row.join(" ").toLowerCase().includes(query));

    if (sortIndex !== null) {
      currentRows.sort((first, second) => compareTableValues(first[sortIndex], second[sortIndex], sortDirection));
    }

    tbody.innerHTML = "";
    currentRows.forEach((row) => {
      const tr = document.createElement("tr");
      tr.innerHTML = row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("");
      tbody.appendChild(tr);
    });
  }

  renderTableRows();

  table.append(thead, tbody);
  tableWrap.appendChild(table);
  wrapper.append(tools, tableWrap);
  return wrapper;
}

function compareTableValues(firstValue, secondValue, direction) {
  const firstNumber = Number(firstValue);
  const secondNumber = Number(secondValue);
  const bothNumeric = Number.isFinite(firstNumber) && Number.isFinite(secondNumber);
  const result = bothNumeric
    ? firstNumber - secondNumber
    : String(firstValue).localeCompare(String(secondValue), "pt-PT", { numeric: true, sensitivity: "base" });

  return direction === "asc" ? result : -result;
}

function createEmptyMessage(message) {
  const paragraph = document.createElement("p");
  paragraph.className = "form-note";
  paragraph.textContent = message;
  return paragraph;
}

async function updateAdminResults(preloadedStudents = null, preloadedChoices = null) {
  if (!adminResults) {
    return;
  }

  adminResults.textContent = "A carregar resultados...";

  try {
    const studentsList = preloadedStudents || adminStudentsCache;
    const choices = preloadedChoices || await choiceRepository.getAll();
    const choicesByAluno = new Map(choices.map((choice) => [String(choice.aluno_id), choice]));
    const rows = studentsList.map((student) => ({
      student,
      choice: choicesByAluno.get(String(student.aluno_id)) || null
    }));

    if (rows.length === 0) {
      adminResults.textContent = "Não existem alunos para os filtros selecionados. Ajusta os filtros ou clica em Limpar filtros para voltar à lista completa.";
      return;
    }

    const wrapper = document.createElement("div");
    const adminActionMessage = document.createElement("p");
    adminActionMessage.className = "message hidden";
    adminActionMessage.setAttribute("role", "status");
    adminActionMessage.setAttribute("aria-live", "polite");

    const bulkActions = document.createElement("div");
    bulkActions.className = "admin-bulk-actions";

    const selectAllLabel = document.createElement("label");
    selectAllLabel.className = "bulk-select";
    selectAllLabel.innerHTML = `<input id="bulk-select-results" type="checkbox"> Selecionar alunos visíveis`;

    const lockSelectedButton = document.createElement("button");
    lockSelectedButton.type = "button";
    lockSelectedButton.className = "secondary danger";
    lockSelectedButton.textContent = "Bloquear selecionados";

    const unlockSelectedButton = document.createElement("button");
    unlockSelectedButton.type = "button";
    unlockSelectedButton.className = "secondary";
    unlockSelectedButton.textContent = "Desbloquear selecionados";

    const markNotRenewingButton = document.createElement("button");
    markNotRenewingButton.type = "button";
    markNotRenewingButton.className = "secondary";
    markNotRenewingButton.textContent = "Marcar como não renovam";

    const reactivateStudentsButton = document.createElement("button");
    reactivateStudentsButton.type = "button";
    reactivateStudentsButton.className = "secondary";
    reactivateStudentsButton.textContent = "Reativar alunos";

    bulkActions.append(selectAllLabel, lockSelectedButton, unlockSelectedButton, markNotRenewingButton, reactivateStudentsButton);

    const table = document.createElement("table");
    table.className = "results-table";
    const thead = document.createElement("thead");
    const tbody = document.createElement("tbody");
    let sortedRows = [...rows];
    let resultSortKey = null;
    let resultSortDirection = "asc";
    const resultHeaders = [
      ["Selecionar", null],
      ["Aluno", "student"],
      ["Turma", "className"],
      ["Curso", "course"],
      ["Situação", "submission"],
      ["Estado", "status"],
      ["Ações", null]
    ];
    const headerRow = document.createElement("tr");

    resultHeaders.forEach(([label, key]) => {
      const th = document.createElement("th");

      if (!key) {
        th.textContent = label;
        headerRow.appendChild(th);
        return;
      }

      const button = document.createElement("button");
      button.type = "button";
      button.className = "table-sort";
      button.textContent = label;
      button.setAttribute("aria-label", `Ordenar por ${label}`);
      button.addEventListener("click", () => {
        if (resultSortKey === key) {
          resultSortDirection = resultSortDirection === "asc" ? "desc" : "asc";
        } else {
          resultSortKey = key;
          resultSortDirection = "asc";
        }

        renderAdminResultRows();
      });
      th.appendChild(button);
      headerRow.appendChild(th);
    });

    thead.appendChild(headerRow);
    table.append(thead, tbody);

    function getAdminResultSortValue(row, key) {
      if (key === "student") {
        return `${row.student.nome} ${row.student.aluno_id}`;
      }

      if (key === "className") {
        return row.student.turma;
      }

      if (key === "course") {
        return row.choice ? getChoiceCourseLabel(row.choice) : getCourseLabel(row.student.curso);
      }

      if (key === "submission") {
        return isStudentNotRenewing(row.student.aluno_id) ? "Não renova" : row.choice ? "Preencheu" : "Por preencher";
      }

      if (key === "status") {
        return row.choice ? (row.choice.estado === "bloqueada" ? "Bloqueada" : "Editável") : "";
      }

      return "";
    }

    function renderAdminResultRows() {
      sortedRows = [...rows];

      if (resultSortKey) {
        sortedRows.sort((first, second) =>
          compareTableValues(
            getAdminResultSortValue(first, resultSortKey),
            getAdminResultSortValue(second, resultSortKey),
            resultSortDirection
          )
        );
      }

      tbody.innerHTML = "";
      sortedRows.forEach(({ student, choice }) => {
      const row = document.createElement("tr");
      const submitted = Boolean(choice);
      const notRenewing = isStudentNotRenewing(student.aluno_id);
      const isLocked = choice?.estado === "bloqueada";
      const rowCourse = choice ? getChoiceCourseLabel(choice) : getCourseLabel(student.curso);
      const situationLabel = notRenewing ? "Não renova" : submitted ? "Preencheu" : "Por preencher";
      const situationClass = notRenewing ? "not-renewing" : submitted ? "editable" : "pending";
      row.innerHTML = `
        <td></td>
        <td>${escapeHtml(student.nome)}<br><small>${escapeHtml(student.aluno_id)}</small></td>
        <td>${escapeHtml(student.turma)}</td>
        <td>${escapeHtml(rowCourse)}</td>
        <td><span class="status-pill ${situationClass}">${situationLabel}</span></td>
        <td>${submitted && !notRenewing ? `<span class="status-pill ${isLocked ? "locked" : "editable"}">${isLocked ? "Bloqueada" : "Editável"}</span>` : "-"}</td>
        <td></td>
      `;

      const selectCell = row.querySelector("td:first-child");
      const actionCell = row.querySelector("td:last-child");

      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.className = "admin-result-checkbox";
      checkbox.value = String(student.aluno_id);
      checkbox.setAttribute("aria-label", `Selecionar ${student.nome}`);
      selectCell.appendChild(checkbox);

      if (notRenewing) {
        const reactivateButton = document.createElement("button");
        reactivateButton.type = "button";
        reactivateButton.className = "secondary";
        reactivateButton.textContent = "Reativar";
        reactivateButton.addEventListener("click", async () => {
          await updateStudentProcessStatus(student.aluno_id, "ativo", reactivateButton, adminActionMessage);
        });
        actionCell.appendChild(reactivateButton);
        tbody.appendChild(row);
        return;
      }

      const button = document.createElement("button");
      button.type = "button";
      button.className = isLocked ? "secondary" : "secondary danger";
      button.textContent = isLocked ? "Desbloquear" : "Bloquear";
      button.classList.toggle("hidden", !submitted);
      button.addEventListener("click", async () => {
        button.disabled = true;

        try {
          await choiceRepository.updateStatus(choice.aluno_id, isLocked ? "submetida" : "bloqueada");
          await updateAdminDashboard();
          await updateCsvOutput();
        } catch (error) {
          adminResults.textContent = error.message;
        }
      });

      const notRenewingButton = document.createElement("button");
      notRenewingButton.type = "button";
      notRenewingButton.className = "secondary";
      notRenewingButton.textContent = "Não renova";
      notRenewingButton.addEventListener("click", async () => {
        await updateStudentProcessStatus(student.aluno_id, "nao_renova", notRenewingButton, adminActionMessage);
      });

      if (submitted) {
        actionCell.appendChild(button);
      }

      actionCell.appendChild(notRenewingButton);
      tbody.appendChild(row);
    });
    }

    renderAdminResultRows();

    const submittedCheckboxes = () => Array.from(table.querySelectorAll(".admin-result-checkbox"));
    const selectedAlunoIds = () => submittedCheckboxes().filter((checkbox) => checkbox.checked).map((checkbox) => checkbox.value);
    const selectedSubmittedAlunoIds = () => selectedAlunoIds().filter((alunoId) => choicesByAluno.has(String(alunoId)) && !isStudentNotRenewing(alunoId));
    const setBulkButtonsState = () => {
      const hasSelection = selectedAlunoIds().length > 0;
      const hasSubmittedSelection = selectedSubmittedAlunoIds().length > 0;
      selectAllCheckbox.disabled = submittedCheckboxes().length === 0;
      lockSelectedButton.disabled = !hasSubmittedSelection;
      unlockSelectedButton.disabled = !hasSubmittedSelection;
      markNotRenewingButton.disabled = !hasSelection;
      reactivateStudentsButton.disabled = !hasSelection;
    };
    const selectAllCheckbox = selectAllLabel.querySelector("input");
    selectAllCheckbox.addEventListener("change", () => {
      submittedCheckboxes().forEach((checkbox) => {
        checkbox.checked = selectAllCheckbox.checked;
      });
      setBulkButtonsState();
    });
    table.addEventListener("change", (event) => {
      if (!event.target.classList.contains("admin-result-checkbox")) {
        return;
      }

      const checkboxes = submittedCheckboxes();
      selectAllCheckbox.checked = checkboxes.length > 0 && checkboxes.every((checkbox) => checkbox.checked);
      setBulkButtonsState();
    });

    lockSelectedButton.addEventListener("click", () => updateSelectedChoicesStatus(selectedSubmittedAlunoIds(), "bloqueada"));
    unlockSelectedButton.addEventListener("click", () => updateSelectedChoicesStatus(selectedSubmittedAlunoIds(), "submetida"));
    markNotRenewingButton.addEventListener("click", () => updateSelectedStudentsProcessStatus(selectedAlunoIds(), "nao_renova", markNotRenewingButton, adminActionMessage));
    reactivateStudentsButton.addEventListener("click", () => updateSelectedStudentsProcessStatus(selectedAlunoIds(), "ativo", reactivateStudentsButton, adminActionMessage));
    setBulkButtonsState();

    adminResults.innerHTML = "";
    wrapper.append(adminActionMessage, bulkActions, table);
    adminResults.appendChild(wrapper);
  } catch (error) {
    adminResults.textContent = error.message;
  }
}

async function updateSelectedChoicesStatus(alunoIds, estado) {
  if (alunoIds.length === 0) {
    return;
  }

  adminResults.setAttribute("aria-busy", "true");

  try {
    await Promise.all(alunoIds.map((alunoId) => choiceRepository.updateStatus(alunoId, estado)));
    await updateAdminDashboard();
    await updateCsvOutput();
  } catch (error) {
    adminResults.textContent = error.message;
  } finally {
    adminResults.removeAttribute("aria-busy");
  }
}

async function updateStudentProcessStatus(alunoId, estado, triggerButton = null, messageElement = null) {
  await updateSelectedStudentsProcessStatus([alunoId], estado, triggerButton, messageElement);
}

async function updateSelectedStudentsProcessStatus(alunoIds, estado, triggerButton = null, messageElement = null) {
  if (alunoIds.length === 0) {
    return;
  }

  const originalText = triggerButton?.textContent;
  adminResults.setAttribute("aria-busy", "true");

  if (triggerButton) {
    triggerButton.disabled = true;
    triggerButton.textContent = "A guardar...";
  }

  if (messageElement) {
    messageElement.className = "message";
    messageElement.textContent = "A guardar alteração...";
  }

  try {
    await Promise.all(alunoIds.map((alunoId) => studentStatusRepository.updateStatus(alunoId, estado)));
    alunoIds.forEach((alunoId) => {
      adminStudentStatusesCache.set(String(alunoId), estado);
    });
    await updateAdminDashboard();

    if (messageElement) {
      messageElement.className = "message success";
      messageElement.textContent = estado === "nao_renova"
        ? "Aluno(s) marcado(s) como não renovam."
        : "Aluno(s) reativado(s).";
    }
    await updateCsvOutput();
  } catch (error) {
    const message = error.message || "Não foi possível guardar a alteração.";

    if (messageElement) {
      messageElement.className = "message warning";
      messageElement.textContent = message;
    }

    window.alert(message);
  } finally {
    if (triggerButton) {
      triggerButton.disabled = false;
      triggerButton.textContent = originalText;
    }

    adminResults.removeAttribute("aria-busy");
  }
}

function buildChoicesCsv(choices) {
  if (choices.length === 0) {
    return "";
  }

  const rows = [
    [
      "aluno_id",
      "nome",
      "turma",
      "curso",
      "processo_id",
      "ano",
      "autenticado_com",
      "prioridade_1_disciplina_1",
      "prioridade_1_disciplina_2",
      "prioridade_2_disciplina_1",
      "prioridade_2_disciplina_2",
      "prioridade_3_disciplina_1",
      "prioridade_3_disciplina_2",
      "faz_maths_outubro",
      "cambridge_psychology",
      "cambridge_economics",
      "cambridge_disciplina_extra",
      "curso_10",
      "cambridge_10",
      "tipo_cambridge",
      "estado_cambridge",
      "disciplinas_automaticas_10",
      "disciplina_opcao_10",
      "submetido_em",
      "estado"
    ],
    ...choices.map((choice) => {
      const prioritySubjects = getPrioritySubjectsForExport(choice);

      return [
        choice.aluno_id,
        choice.nome,
        choice.turma,
        choice.curso,
        getProcessId(choice),
        getProcessYear(choice),
        choice.autenticado_com || "",
        ...prioritySubjects,
        choice.cambridge?.faz_maths_outubro ?? "",
        choice.cambridge?.psychology ?? "",
        choice.cambridge?.economics ?? "",
        choice.cambridge?.disciplina_extra || "",
        choice.escolha_10?.curso_label || "",
        choice.escolha_10?.cambridge ?? "",
        choice.escolha_10?.tipo_cambridge || "",
        choice.escolha_10?.estado_cambridge || "",
        (choice.escolha_10?.disciplinas_automaticas || []).join(" | "),
        choice.escolha_10?.disciplina_opcao || "",
        choice.submetido_em,
        choice.estado || "submetida"
      ];
    })
  ];

  return rows.map((row) => row.map(escapeCsvValue).join(",")).join("\n");
}

function getPrioritySubjectsForExport(choice) {
  if (choice.escolha_10) {
    return getTenthGradeSubjectsForExport(choice.escolha_10);
  }

  if (choice.cambridge) {
    return getCambridgeSubjectsForExport(choice.cambridge);
  }

  if (Array.isArray(choice.prioridades)) {
    const subjects = choice.prioridades.flatMap((item) => item.subjects || []);
    return [...subjects, "", "", "", "", "", ""].slice(0, 6);
  }

  // Compatibilidade com escolhas criadas antes do sistema de prioridades.
  if (Array.isArray(choice.disciplinas)) {
    return [choice.disciplinas[0] || "", choice.disciplinas[1] || "", "", "", "", ""];
  }

  return ["", "", "", "", "", ""];
}

function getTenthGradeSubjectsForExport(tenthGrade) {
  const subjects = [
    ...(tenthGrade.disciplinas_automaticas || []),
    tenthGrade.disciplina_opcao
  ].filter(Boolean);

  return [...subjects, "", "", "", "", "", ""].slice(0, 6);
}

function getCambridgeSubjectsForExport(cambridge) {
  const subjects = [];

  if (cambridge.psychology) {
    subjects.push("Psychology");
  }

  if (cambridge.economics) {
    subjects.push("Economics");
  }

  if (cambridge.disciplina_extra) {
    subjects.push(cambridge.disciplina_extra);
  }

  return [...subjects, "", "", "", "", "", ""].slice(0, 6);
}

function escapeCsvValue(value) {
  const text = String(value || "");

  if (/[",\n]/.test(text)) {
    return `"${text.replaceAll('"', '""')}"`;
  }

  return text;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function initProcessUi() {
  document.title = CURRENT_PROCESS.title;

  if (pageTitle) {
    pageTitle.textContent = CURRENT_PROCESS.title;
  }
}

initProcessUi();
authInitPromise = initAuth();
