const DATA_SOURCE = "alunos.csv";
const STORAGE_KEY = "escolhas12ano.resultados";
const ADMIN_EMAIL = "mat_tic_josesantos@c-guadalupe.com";
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
    requiredGroupName: "disciplinas obrigatórias",
    required: ["Oficina de Artes", "Oficina de Multimédia B"],
    optional: [],
    exactSubjects: ["Oficina de Artes", "Oficina de Multimédia B"],
    ruleText: "As únicas disciplinas possíveis são Oficina de Artes e Oficina de Multimédia B. Como só existe um par válido, não é possível criar 3 prioridades diferentes sem uma regra administrativa especial."
  }
};

const loginMessage = document.querySelector("#login-message");
const authWarning = document.querySelector("#auth-warning");
const signInButton = document.querySelector("#sign-in");
const headerSession = document.querySelector("#header-session");
const headerAuthEmail = document.querySelector("#header-auth-email");
const profilePhoto = document.querySelector("#profile-photo");
const headerSignOutButton = document.querySelector("#header-sign-out");
const homeLink = document.querySelector("#home-link");
const adminNav = document.querySelector("#admin-nav");
const studentArea = document.querySelector("#student-area");
const summaryArea = document.querySelector("#summary-area");
const summaryMessage = document.querySelector("#summary-message");
const summaryTable = document.querySelector("#summary-table");
const goHomeButton = document.querySelector("#go-home");
const editChoiceButton = document.querySelector("#edit-choice");
const choiceStatus = document.querySelector("#choice-status");
const choicesForm = document.querySelector("#choices-form");
const prioritiesList = document.querySelector("#priorities-list");
const validationMessage = document.querySelector("#validation-message");
const submitChoiceButton = document.querySelector("#submit-choice");
const confirmation = document.querySelector("#confirmation");
const csvOutput = document.querySelector("#csv-output");
const adminResults = document.querySelector("#admin-results");
const refreshResultsButton = document.querySelector("#refresh-results");
const exportCsvButton = document.querySelector("#export-csv");
const downloadCsvButton = document.querySelector("#download-csv");
const clearResultsButton = document.querySelector("#clear-results");

let students = [];
let currentStudent = null;
let currentChoice = null;
let msalClient = null;
let signedInAccount = null;
let authInitPromise = null;
let supabaseClient = null;
let supabaseInitPromise = null;

const loginRequest = {
  scopes: ["openid", "profile", "email", "User.Read"]
};

const studentRepository = {
  async getAll() {
    const client = await ensureSupabaseReady();

    if (client) {
      const { data, error } = await client
        .from("alunos")
        .select("aluno_id,nome,turma,curso,email")
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
        .select("aluno_id,nome,turma,curso,email")
        .eq("aluno_id", studentId)
        .maybeSingle();

      if (error) {
        throw new Error("Não foi possível consultar a base de dados de alunos.");
      }

      return data ? normalizeStudent(data) : null;
    }

    if (students.length === 0) {
      students = await this.getAll();
    }

    return students.find((student) => student.aluno_id === studentId);
  },

  async findByEmail(email) {
    const normalizedEmail = email.toLowerCase();
    const client = await ensureSupabaseReady();

    if (client) {
      const { data, error } = await client
        .from("alunos")
        .select("aluno_id,nome,turma,curso,email")
        .eq("email", normalizedEmail)
        .maybeSingle();

      if (error) {
        throw new Error("Não foi possível validar os dados do aluno.");
      }

      return data ? normalizeStudent(data) : null;
    }

    if (students.length === 0) {
      students = await this.getAll();
    }

    return students.find((student) => (student.email || "").toLowerCase() === normalizedEmail) || null;
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
        .maybeSingle();

      if (error) {
        throw new Error("Não foi possível consultar a submissão existente.");
      }

      return data ? mapChoiceFromDatabase(data) : null;
    }

    const choices = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
    return choices.find((choice) => choice.aluno_id === String(alunoId)) || null;
  },

  async getAll() {
    const client = await ensureSupabaseReady();

    if (client) {
      const { data, error } = await client
        .from("escolhas")
        .select("*")
        .order("submetido_em", { ascending: false });

      if (error) {
        throw new Error("Não foi possível carregar as escolhas guardadas.");
      }

      return data.map(mapChoiceFromDatabase);
    }

    return JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
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
        ? client.from("escolhas").update(payload).eq("aluno_id", String(choice.aluno_id))
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
      .eq("aluno_id", String(alunoId));

    if (error) {
      throw new Error("Não foi possível atualizar o estado da submissão.");
    }
  },

  async clear() {
    const client = await ensureSupabaseReady();

    if (client) {
      const { error } = await client.from("escolhas").delete().neq("aluno_id", "__nunca__");

      if (error) {
        throw new Error("Não foi possível limpar as escolhas na base de dados.");
      }

      return;
    }

    localStorage.removeItem(STORAGE_KEY);
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

homeLink.addEventListener("click", (event) => {
  event.preventDefault();
  showHome();
});

goHomeButton.addEventListener("click", showHome);

editChoiceButton.addEventListener("click", () => {
  if (!currentStudent) {
    return;
  }

  showChoiceEditor();
});

choicesForm.addEventListener("change", () => {
  enforcePriorityLimits();
  updateValidation();
});

choicesForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const priorities = getSelectedPriorities();
  const validation = validatePriorities(priorities, currentStudent.curso);

  if (!validation.valid) {
    updateValidation();
    return;
  }

  const choice = {
    aluno_id: currentStudent.aluno_id,
    nome: currentStudent.nome,
    turma: currentStudent.turma,
    curso: currentStudent.curso,
    autenticado_com: getSignedInEmail(),
    prioridades: priorities,
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

refreshResultsButton.addEventListener("click", updateAdminResults);
exportCsvButton.addEventListener("click", updateCsvOutput);

downloadCsvButton.addEventListener("click", async () => {
  const csv = buildChoicesCsv(await choiceRepository.getAll());
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "escolhas-12ano.csv";
  link.click();
  URL.revokeObjectURL(link.href);
});

clearResultsButton.addEventListener("click", async () => {
  const confirmed = window.confirm("Queres limpar todos os resultados guardados na base de dados?");

  if (confirmed) {
    try {
      await choiceRepository.clear();
      await updateCsvOutput();
      await updateAdminResults();
    } catch (error) {
      csvOutput.value = error.message;
    }
  }
});

async function loadSignedInStudent() {
  if (!signedInAccount || getSignedInEmail() === ADMIN_EMAIL) {
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

    if (!courseRules[student.curso]) {
      showLoginError("O teu curso não está configurado nesta aplicação. Contacta a administração.");
      return;
    }

    currentStudent = student;
    currentChoice = await choiceRepository.getByAlunoId(student.aluno_id);
    loginMessage.textContent = "";
    document.querySelector("#entrada").classList.add("hidden");
    renderStudentArea(student, currentChoice);

    if (currentChoice) {
      showSummaryPage(currentChoice);
    } else {
      showChoiceEditor();
    }
  } catch (error) {
    showLoginError(error.message);
  }
}

function renderStudentArea(student, existingChoice = null) {
  const rules = courseRules[student.curso];

  document.querySelector("#student-name").textContent = student.nome;
  document.querySelector("#student-class").textContent = student.turma;
  document.querySelector("#student-course").textContent = rules.label;
  document.querySelector("#course-rules").textContent = rules.ruleText;

  prioritiesList.innerHTML = "";

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
    enforcePriorityLimits();
    return;
  }

  document.querySelectorAll('input[type="checkbox"][name^="priority-"]').forEach((input) => {
    input.disabled = true;
  });

  submitChoiceButton.disabled = true;
}

function showHome() {
  studentArea.classList.add("hidden");
  summaryArea.classList.add("hidden");

  if (signedInAccount && getSignedInEmail() !== ADMIN_EMAIL) {
    document.querySelector("#entrada").classList.remove("hidden");
  }

  window.history.replaceState(null, "", HOME_URL);
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function showChoiceEditor() {
  if (!currentStudent) {
    return;
  }

  document.querySelector("#entrada").classList.add("hidden");
  summaryArea.classList.add("hidden");
  studentArea.classList.remove("hidden");
  setChoicesLocked(currentChoice?.estado === "bloqueada");
  updateValidation();
  window.history.replaceState(null, "", "#opcoes");
  studentArea.scrollIntoView({ behavior: "smooth", block: "start" });
}

function showSummaryPage(choice) {
  studentArea.classList.add("hidden");
  document.querySelector("#entrada").classList.add("hidden");
  summaryArea.classList.remove("hidden");
  renderSummaryTable(choice);

  const isLocked = choice?.estado === "bloqueada";
  summaryMessage.textContent = isLocked
    ? "A submissão está bloqueada pela administração e não pode ser alterada."
    : "A submissão está guardada e ainda pode ser alterada.";
  editChoiceButton.classList.toggle("hidden", isLocked);

  window.history.replaceState(null, "", "#resumo");
  summaryArea.scrollIntoView({ behavior: "smooth", block: "start" });
}

function renderSummaryTable(choice) {
  if (!choice) {
    summaryTable.textContent = "Ainda não existe submissão.";
    return;
  }

  const table = document.createElement("table");
  table.className = "results-table";
  table.innerHTML = `
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

  choice.prioridades.forEach((item) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>Prioridade ${item.priority}</td>
      <td>${escapeHtml(item.subjects[0])}</td>
      <td>${escapeHtml(item.subjects[1])}</td>
    `;
    tbody.appendChild(row);
  });

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

  const validation = validatePriorities(getSelectedPriorities(), currentStudent.curso);
  validationMessage.textContent = validation.message;
  validationMessage.className = `validation ${validation.valid ? "valid" : "invalid"}`;
  submitChoiceButton.disabled = !validation.valid;
}

function validatePriorities(priorities, courseKey) {
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

function findOverusedRequiredSubject(priorities, courseKey) {
  const rules = courseRules[courseKey];
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
  const rules = courseRules[courseKey];

  if (selectedSubjects.length !== 2) {
    return {
      valid: false,
      detail: `seleciona exatamente duas disciplinas. Tens ${selectedSubjects.length} selecionada(s).`
    };
  }

  if (courseKey === "ArtesVisuais") {
    const hasOnlyExactSubjects = rules.exactSubjects.every((subject) => selectedSubjects.includes(subject));

    return hasOnlyExactSubjects
      ? { valid: true }
      : { valid: false, detail: "em Artes Visuais tens de escolher Oficina de Artes e Oficina de Multimédia B." };
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
    email: student.email || ""
  };
}

function mapChoiceToDatabase(choice, existingChoice = null) {
  const prioritySubjects = getPrioritySubjectsForExport(choice);
  const now = new Date().toISOString();

  return {
    aluno_id: String(choice.aluno_id),
    nome: choice.nome,
    turma: choice.turma,
    curso: choice.curso,
    email_autenticado: choice.autenticado_com,
    prioridade_1_disciplina_1: prioritySubjects[0],
    prioridade_1_disciplina_2: prioritySubjects[1],
    prioridade_2_disciplina_1: prioritySubjects[2],
    prioridade_2_disciplina_2: prioritySubjects[3],
    prioridade_3_disciplina_1: prioritySubjects[4],
    prioridade_3_disciplina_2: prioritySubjects[5],
    submetido_em: existingChoice?.submetido_em || choice.submetido_em,
    atualizado_em: existingChoice ? now : null,
    estado: existingChoice?.estado || "submetida"
  };
}

function mapChoiceFromDatabase(row) {
  return {
    aluno_id: row.aluno_id,
    nome: row.nome,
    turma: row.turma,
    curso: row.curso,
    autenticado_com: row.email_autenticado,
    prioridades: [
      {
        priority: 1,
        subjects: [row.prioridade_1_disciplina_1, row.prioridade_1_disciplina_2]
      },
      {
        priority: 2,
        subjects: [row.prioridade_2_disciplina_1, row.prioridade_2_disciplina_2]
      },
      {
        priority: 3,
        subjects: [row.prioridade_3_disciplina_1, row.prioridade_3_disciplina_2]
      }
    ],
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

function updateAuthUi() {
  const isSignedIn = Boolean(signedInAccount);
  const isAdmin = isSignedIn && getSignedInEmail() === ADMIN_EMAIL;

  document.querySelector("#autenticacao").classList.toggle("hidden", isSignedIn);
  document.querySelectorAll(".requires-auth").forEach((element) => {
    element.classList.toggle("hidden", !isSignedIn || isAdmin);
  });

  document.querySelectorAll(".admin-only").forEach((element) => {
    element.classList.toggle("hidden", !isAdmin);
  });

  adminNav.classList.toggle("hidden", !isAdmin);
  headerSession.classList.toggle("hidden", !isSignedIn);
  signInButton.disabled = false;

  if (isSignedIn) {
    headerAuthEmail.textContent = getSignedInDisplayName();
    loadProfilePhoto();

    if (isAdmin) {
      studentArea.classList.add("hidden");
      updateCsvOutput();
      updateAdminResults();
    } else {
      loadSignedInStudent();
    }
  } else {
    headerAuthEmail.textContent = "";
    profilePhoto.classList.add("hidden");
    profilePhoto.removeAttribute("src");
    studentArea.classList.add("hidden");
    summaryArea.classList.add("hidden");
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
    const choices = await choiceRepository.getAll();
    const csv = buildChoicesCsv(choices);
    csvOutput.value = csv || "Sem resultados submetidos.";
  } catch (error) {
    csvOutput.value = error.message;
  }
}

async function updateAdminResults() {
  if (!adminResults) {
    return;
  }

  adminResults.textContent = "A carregar resultados...";

  try {
    const choices = await choiceRepository.getAll();

    if (choices.length === 0) {
      adminResults.textContent = "Ainda não existem submissões.";
      return;
    }

    const table = document.createElement("table");
    table.className = "results-table";
    table.innerHTML = `
      <thead>
        <tr>
          <th>Aluno</th>
          <th>Turma</th>
          <th>Curso</th>
          <th>Estado</th>
          <th>Ações</th>
        </tr>
      </thead>
      <tbody></tbody>
    `;

    const tbody = table.querySelector("tbody");

    choices.forEach((choice) => {
      const row = document.createElement("tr");
      const isLocked = choice.estado === "bloqueada";
      row.innerHTML = `
        <td>${escapeHtml(choice.nome)}<br><small>${escapeHtml(choice.aluno_id)}</small></td>
        <td>${escapeHtml(choice.turma)}</td>
        <td>${escapeHtml(choice.curso)}</td>
        <td><span class="status-pill ${isLocked ? "locked" : "editable"}">${isLocked ? "Bloqueada" : "Editável"}</span></td>
        <td></td>
      `;

      const actionCell = row.querySelector("td:last-child");
      const button = document.createElement("button");
      button.type = "button";
      button.className = isLocked ? "secondary" : "secondary danger";
      button.textContent = isLocked ? "Desbloquear" : "Bloquear";
      button.addEventListener("click", async () => {
        button.disabled = true;

        try {
          await choiceRepository.updateStatus(choice.aluno_id, isLocked ? "submetida" : "bloqueada");
          await updateAdminResults();
          await updateCsvOutput();
        } catch (error) {
          adminResults.textContent = error.message;
        }
      });

      actionCell.appendChild(button);
      tbody.appendChild(row);
    });

    adminResults.innerHTML = "";
    adminResults.appendChild(table);
  } catch (error) {
    adminResults.textContent = error.message;
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
      "autenticado_com",
      "prioridade_1_disciplina_1",
      "prioridade_1_disciplina_2",
      "prioridade_2_disciplina_1",
      "prioridade_2_disciplina_2",
      "prioridade_3_disciplina_1",
      "prioridade_3_disciplina_2",
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
        choice.autenticado_com || "",
        ...prioritySubjects,
        choice.submetido_em,
        choice.estado || "submetida"
      ];
    })
  ];

  return rows.map((row) => row.map(escapeCsvValue).join(",")).join("\n");
}

function getPrioritySubjectsForExport(choice) {
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

function escapeCsvValue(value) {
  const text = String(value || "");

  if (/[",\n]/.test(text)) {
    return `"${text.replaceAll('"', '""')}"`;
  }

  return text;
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

authInitPromise = initAuth();
