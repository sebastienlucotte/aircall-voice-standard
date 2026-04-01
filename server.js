require("dotenv").config();

const express = require("express");
const nodemailer = require("nodemailer");
const { google } = require("googleapis");

const app = express();
app.use(express.json());

const PORT = process.env.PORT || 3000;
const API_BEARER_TOKEN = process.env.API_BEARER_TOKEN || "change-me";

// SMTP Gmail
const SMTP_HOST = "smtp.gmail.com";
const SMTP_PORT = 465;
const SMTP_SECURE = true;
const SMTP_USER = "appel.rubiomonocoat@gmail.com";
const SMTP_PASS = process.env.SMTP_PASS;
const MAIL_FROM = "appel.rubiomonocoat@gmail.com";

// Google Sheets
const SHEET_ID = process.env.GOOGLE_SHEETS_SPREADSHEET_ID;
const GOOGLE_SERVICE_ACCOUNT_EMAIL = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
const GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY =
  process.env.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY?.replace(/\\n/g, "\n");
const SHEET_NAME = "Logs";

// Contacts de routage
const CONTACTS = {
  baptiste: {
    id: "baptiste",
    name: "Baptiste Verriele",
    email: "baptiste@rubiomonocoat.fr",
    targetType: "external",
    targetValue: "+33675859240",
  },
  guillaume: {
    id: "guillaume",
    name: "Guillaume Nepveu",
    email: "guillaume@rubiomonocoat.fr",
    targetType: "external",
    targetValue: "+33607122212",
  },
  laurent: {
    id: "laurent",
    name: "Laurent Moreau",
    email: "laurent@rubiomonocoat.fr",
    targetType: "external",
    targetValue: "+33608660394",
  },
  antony: {
    id: "antony",
    name: "Antony Grasser",
    email: "antony@rubiomonocoat.fr",
    targetType: "external",
    targetValue: "+33698281840",
  },
  benjamin: {
    id: "benjamin",
    name: "Benjamin Hardial",
    email: "benjamin@rubiomonocoat.fr",
    targetType: "external",
    targetValue: "+33786358881",
  },
  particuliers: {
    id: "particuliers",
    name: "Ligne particuliers",
    email: "sebastien@rubiomonocoat.fr",
    targetType: "external",
    targetValue: "+33757941786",
  },
  sebastien: {
    id: "sebastien",
    name: "Sébastien",
    email: "sebastien@rubiomonocoat.fr",
    targetType: "external",
    targetValue: "+33621414949",
  },
};

// Routage par département
const ROUTING = {
  // Nord / IDF
  "02": CONTACTS.baptiste,
  "08": CONTACTS.baptiste,
  "27": CONTACTS.baptiste,
  "51": CONTACTS.baptiste,
  "59": CONTACTS.baptiste,
  "60": CONTACTS.baptiste,
  "62": CONTACTS.baptiste,
  "75": CONTACTS.baptiste,
  "76": CONTACTS.baptiste,
  "77": CONTACTS.baptiste,
  "78": CONTACTS.baptiste,
  "80": CONTACTS.baptiste,
  "91": CONTACTS.baptiste,
  "92": CONTACTS.baptiste,
  "93": CONTACTS.baptiste,
  "94": CONTACTS.baptiste,
  "95": CONTACTS.baptiste,

  // Ouest
  "03": CONTACTS.guillaume,
  "14": CONTACTS.guillaume,
  "18": CONTACTS.guillaume,
  "22": CONTACTS.guillaume,
  "28": CONTACTS.guillaume,
  "29": CONTACTS.guillaume,
  "35": CONTACTS.guillaume,
  "37": CONTACTS.guillaume,
  "41": CONTACTS.guillaume,
  "44": CONTACTS.guillaume,
  "45": CONTACTS.guillaume,
  "49": CONTACTS.guillaume,
  "50": CONTACTS.guillaume,
  "53": CONTACTS.guillaume,
  "56": CONTACTS.guillaume,
  "58": CONTACTS.guillaume,
  "61": CONTACTS.guillaume,
  "72": CONTACTS.guillaume,
  "85": CONTACTS.guillaume,
  "89": CONTACTS.guillaume,

  // Sud-Ouest
  "09": CONTACTS.laurent,
  "16": CONTACTS.laurent,
  "17": CONTACTS.laurent,
  "19": CONTACTS.laurent,
  "23": CONTACTS.laurent,
  "24": CONTACTS.laurent,
  "31": CONTACTS.laurent,
  "32": CONTACTS.laurent,
  "33": CONTACTS.laurent,
  "36": CONTACTS.laurent,
  "40": CONTACTS.laurent,
  "46": CONTACTS.laurent,
  "47": CONTACTS.laurent,
  "64": CONTACTS.laurent,
  "65": CONTACTS.laurent,
  "79": CONTACTS.laurent,
  "81": CONTACTS.laurent,
  "82": CONTACTS.laurent,
  "86": CONTACTS.laurent,
  "87": CONTACTS.laurent,

  // Est
  "01": CONTACTS.antony,
  "10": CONTACTS.antony,
  "21": CONTACTS.antony,
  "25": CONTACTS.antony,
  "39": CONTACTS.antony,
  "52": CONTACTS.antony,
  "54": CONTACTS.antony,
  "55": CONTACTS.antony,
  "57": CONTACTS.antony,
  "67": CONTACTS.antony,
  "68": CONTACTS.antony,
  "69": CONTACTS.antony,
  "70": CONTACTS.antony,
  "71": CONTACTS.antony,
  "73": CONTACTS.antony,
  "74": CONTACTS.antony,
  "88": CONTACTS.antony,
  "90": CONTACTS.antony,

  // Sud-Est
  "04": CONTACTS.benjamin,
  "05": CONTACTS.benjamin,
  "06": CONTACTS.benjamin,
  "07": CONTACTS.benjamin,
  "11": CONTACTS.benjamin,
  "12": CONTACTS.benjamin,
  "13": CONTACTS.benjamin,
  "15": CONTACTS.benjamin,
  "20": CONTACTS.benjamin, // Corse 2A / 2B regroupée
  "26": CONTACTS.benjamin,
  "30": CONTACTS.benjamin,
  "34": CONTACTS.benjamin,
  "38": CONTACTS.benjamin,
  "42": CONTACTS.benjamin,
  "43": CONTACTS.benjamin,
  "48": CONTACTS.benjamin,
  "63": CONTACTS.benjamin,
  "66": CONTACTS.benjamin,
  "83": CONTACTS.benjamin,
  "84": CONTACTS.benjamin,
};

// Correspondance noms de départements -> code
const DEPARTMENT_NAME_TO_CODE = {
  "ain": "01",
  "aisne": "02",
  "allier": "03",
  "alpes de haute provence": "04",
  "alpes-de-haute-provence": "04",
  "hautes alpes": "05",
  "hautes-alpes": "05",
  "alpes maritimes": "06",
  "alpes-maritimes": "06",
  "ardeche": "07",
  "ardennes": "08",
  "ariege": "09",
  "aube": "10",
  "aude": "11",
  "aveyron": "12",
  "bouches du rhone": "13",
  "bouches-du-rhone": "13",
  "bouches du rhône": "13",
  "bouches-du-rhône": "13",
  "calvados": "14",
  "cantal": "15",
  "charente": "16",
  "charente maritime": "17",
  "charente-maritime": "17",
  "cher": "18",
  "correze": "19",
  "corrèze": "19",
  "corse": "20",
  "corse du sud": "2A",
  "corse-du-sud": "2A",
  "haute corse": "2B",
  "haute-corse": "2B",
  "cote d'or": "21",
  "côte d'or": "21",
  "cote-d'or": "21",
  "côtes d'armor": "22",
  "cotes d'armor": "22",
  "cotes d armor": "22",
  "creuse": "23",
  "dordogne": "24",
  "doubs": "25",
  "drome": "26",
  "drôme": "26",
  "eure": "27",
  "eure et loir": "28",
  "eure-et-loir": "28",
  "finistere": "29",
  "finistère": "29",
  "gard": "30",
  "haute garonne": "31",
  "haute-garonne": "31",
  "gers": "32",
  "gironde": "33",
  "herault": "34",
  "hérault": "34",
  "ille et vilaine": "35",
  "ille-et-vilaine": "35",
  "indre": "36",
  "indre et loire": "37",
  "indre-et-loire": "37",
  "isere": "38",
  "isère": "38",
  "jura": "39",
  "landes": "40",
  "loir et cher": "41",
  "loir-et-cher": "41",
  "loire": "42",
  "haute loire": "43",
  "haute-loire": "43",
  "loire atlantique": "44",
  "loire-atlantique": "44",
  "loiret": "45",
  "lot": "46",
  "lot et garonne": "47",
  "lot-et-garonne": "47",
  "lozere": "48",
  "lozère": "48",
  "maine et loire": "49",
  "maine-et-loire": "49",
  "manche": "50",
  "marne": "51",
  "haute marne": "52",
  "haute-marne": "52",
  "mayenne": "53",
  "meurthe et moselle": "54",
  "meurthe-et-moselle": "54",
  "meuse": "55",
  "morbihan": "56",
  "moselle": "57",
  "nievre": "58",
  "nièvre": "58",
  "nord": "59",
  "oise": "60",
  "orne": "61",
  "pas de calais": "62",
  "pas-de-calais": "62",
  "puy de dome": "63",
  "puy-de-dome": "63",
  "puy-de-dôme": "63",
  "pyrenees atlantiques": "64",
  "pyrenees-atlantiques": "64",
  "pyrénées atlantiques": "64",
  "pyrénées-atlantiques": "64",
  "hautes pyrenees": "65",
  "hautes-pyrenees": "65",
  "hautes pyrénées": "65",
  "hautes-pyrénées": "65",
  "pyrenees orientales": "66",
  "pyrenees-orientales": "66",
  "pyrénées orientales": "66",
  "pyrénées-orientales": "66",
  "bas rhin": "67",
  "bas-rhin": "67",
  "haut rhin": "68",
  "haut-rhin": "68",
  "rhone": "69",
  "rhône": "69",
  "haute saone": "70",
  "haute-saone": "70",
  "haute saône": "70",
  "haute-saône": "70",
  "saone et loire": "71",
  "saône et loire": "71",
  "saone-et-loire": "71",
  "saône-et-loire": "71",
  "sarthe": "72",
  "savoie": "73",
  "haute savoie": "74",
  "haute-savoie": "74",
  "paris": "75",
  "seine maritime": "76",
  "seine-maritime": "76",
  "seine et marne": "77",
  "seine-et-marne": "77",
  "yvelines": "78",
  "deux sevres": "79",
  "deux-sevres": "79",
  "deux sèvres": "79",
  "deux-sèvres": "79",
  "somme": "80",
  "tarn": "81",
  "tarn et garonne": "82",
  "tarn-et-garonne": "82",
  "var": "83",
  "vaucluse": "84",
  "vendee": "85",
  "vendée": "85",
  "vienne": "86",
  "haute vienne": "87",
  "haute-vienne": "87",
  "vosges": "88",
  "yonne": "89",
  "territoire de belfort": "90",
  "essonne": "91",
  "hauts de seine": "92",
  "hauts-de-seine": "92",
  "seine saint denis": "93",
  "seine-saint-denis": "93",
  "val de marne": "94",
  "val-de-marne": "94",
  "val d'oise": "95",
  "val d oise": "95",
  "val-d'oise": "95",

  // DROM / TOM usuels
  "guadeloupe": "971",
  "martinique": "972",
  "guyane": "973",
  "la reunion": "974",
  "la réunion": "974",
  "reunion": "974",
  "réunion": "974",
  "mayotte": "976",
  "saint pierre et miquelon": "975",
  "saint-pierre-et-miquelon": "975",
  "saint barthelemy": "977",
  "saint-barthelemy": "977",
  "saint barthélemy": "977",
  "saint-barthélemy": "977",
  "saint martin": "978",
  "saint-martin": "978",
  "wallis et futuna": "986",
  "wallis-et-futuna": "986",
  "polynesie francaise": "987",
  "polynesie-francaise": "987",
  "polynésie française": "987",
  "polynésie-française": "987",
  "nouvelle caledonie": "988",
  "nouvelle-caledonie": "988",
  "nouvelle calédonie": "988",
  "nouvelle-calédonie": "988",
};

function checkAuth(req, res, next) {
  const header = req.headers.authorization || "";

  if (!header.startsWith("Bearer ")) {
    return res.status(401).json({ error: "Missing bearer token" });
  }

  const token = header.slice(7);

  if (token !== API_BEARER_TOKEN) {
    return res.status(401).json({ error: "Invalid bearer token" });
  }

  next();
}

function stripAccents(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function sanitizeText(value) {
  if (value == null) return "";
  return String(value).replace(/\s+/g, " ").trim();
}

function normalizeLooseText(value) {
  return stripAccents(value)
    .toLowerCase()
    .replace(/['’]/g, "'")
    .replace(/[^a-z0-9'\-\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function parseAttempts(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function extractCallerTypeFromText(text) {
  if (!text) return "";

  const raw = normalizeLooseText(text);

  if (raw.includes("distributeur")) return "distributeur";
  if (raw.includes("professionnel")) return "professionnel";
  if (raw.includes("particulier")) return "particulier";

  return "";
}

function extractRequestObjectFromText(text) {
  if (!text) return "";

  const raw = String(text).trim();

  const patterns = [
    /besoin principal[^.:\n]*[.:]?\s*(.+)$/i,
    /objet[^.:\n]*[.:]?\s*(.+)$/i,
    /demande[^.:\n]*[.:]?\s*(.+)$/i,
    /motif[^.:\n]*[.:]?\s*(.+)$/i,
  ];

  for (const pattern of patterns) {
    const match = raw.match(pattern);
    if (match && match[1]) {
      return sanitizeText(match[1]);
    }
  }

  return "";
}

function extractVoiceAgentData(body) {
  const noteRaw =
    body.note ||
    body.notes ||
    body.agentNote ||
    body.voiceAgentNote ||
    body.summary ||
    body.transcript ||
    body.memo ||
    "";

  const sourceAgent =
    body.agentName ||
    body.voiceAgentName ||
    body.source ||
    "AI Voice Agent";

  const callerTypeDirect =
    body.callerType ||
    body.contactType ||
    body.customerType ||
    body.customer_type ||
    "";

  const requestObjectDirect =
    body.requestObject ||
    body.reasonLabel ||
    body.callReason ||
    body.intent ||
    body.request_reason ||
    "";

  const callId =
    body.callId ||
    body.call_id ||
    body.aircallCallId ||
    body.aircall_call_id ||
    "";

  const callUuid =
    body.callUuid ||
    body.call_uuid ||
    body.uuid ||
    "";

  return {
    callerType: sanitizeText(callerTypeDirect) || extractCallerTypeFromText(noteRaw),
    requestObject: sanitizeText(requestObjectDirect) || extractRequestObjectFromText(noteRaw),
    sourceAgent: sanitizeText(sourceAgent),
    callId: sanitizeText(callId),
    callUuid: sanitizeText(callUuid),
    noteRaw: sanitizeText(noteRaw),
  };
}

function findDepartmentCodeByName(input) {
  const normalized = normalizeLooseText(input);

  if (!normalized) return "";

  if (DEPARTMENT_NAME_TO_CODE[normalized]) {
    return DEPARTMENT_NAME_TO_CODE[normalized];
  }

  // petite tolérance si la phrase contient le nom du département
  for (const [name, code] of Object.entries(DEPARTMENT_NAME_TO_CODE)) {
    if (normalized === name || normalized.includes(name)) {
      return code;
    }
  }

  return "";
}

function normalizeDepartmentInput(input) {
  if (input == null) return "";

  const raw = String(input).trim().toUpperCase().replace(/#/g, "");
  const normalizedText = normalizeLooseText(input);

  // Détection par nom
  const byName = findDepartmentCodeByName(normalizedText);
  if (byName) return byName;

  // 2A / 2B explicites
  if (
    raw === "2A" || raw === "2 A" ||
    raw === "2B" || raw === "2 B"
  ) {
    return raw.replace(/\s/g, "");
  }

  const alnum = raw.replace(/[^A-Z0-9]/g, "");

  if (alnum === "2A" || alnum === "2B") return alnum;

  // code postal -> département
  if (/^\d{5}$/.test(alnum)) return alnum.slice(0, 2);

  // codes à 3 chiffres (971, 972...)
  if (/^\d{3}$/.test(alnum)) return alnum;

  // autres chaînes numériques plus longues
  if (/^\d{4,}$/.test(alnum)) return alnum.slice(0, 2);

  // ex: "6" => "06"
  if (/^\d$/.test(alnum)) return "0" + alnum;

  if (/^\d{2}$/.test(alnum)) return alnum;

  return "";
}

function normalizeForRouting(normalizedDept) {
  if (normalizedDept === "2A" || normalizedDept === "2B") {
    return "20";
  }
  if (normalizedDept.startsWith("97") || normalizedDept.startsWith("98")) {
    return normalizedDept.slice(0, 2);
  }
  return normalizedDept;
}

function resolveTarget(rawDepartmentInput, attemptsRaw, callerType) {
  const attempts = parseAttempts(attemptsRaw);

  // Règle métier prioritaire : particulier => ligne dédiée
  if (callerType === "particulier") {
    return {
      contact: CONTACTS.particuliers,
      reason: "PARTICULIER_DIRECT_ROUTING",
      departmentInput: "",
      normalizedDept: "",
      routingCode: "",
      attempts,
    };
  }

  const normalizedDept = normalizeDepartmentInput(rawDepartmentInput);
  const routingCode = normalizeForRouting(normalizedDept);

  if (attempts >= 2) {
    return {
      contact: CONTACTS.sebastien,
      reason: "ATTEMPTS_FALLBACK",
      departmentInput: rawDepartmentInput || "",
      normalizedDept,
      routingCode,
      attempts,
    };
  }

  if (!routingCode) {
    return {
      contact: CONTACTS.sebastien,
      reason: "INVALID_OR_MISSING_CODE",
      departmentInput: rawDepartmentInput || "",
      normalizedDept,
      routingCode,
      attempts,
    };
  }

  if (routingCode === "97" || routingCode === "98") {
    return {
      contact: CONTACTS.sebastien,
      reason: "DOM_ROUTED_TO_SEBASTIEN",
      departmentInput: rawDepartmentInput || "",
      normalizedDept,
      routingCode,
      attempts,
    };
  }

  return {
    contact: ROUTING[routingCode] || CONTACTS.sebastien,
    reason: ROUTING[routingCode] ? "MATCH" : "UNKNOWN_CODE_FALLBACK",
    departmentInput: rawDepartmentInput || "",
    normalizedDept,
    routingCode,
    attempts,
  };
}

function getTransporter() {
  if (!SMTP_USER || !SMTP_PASS) {
    return null;
  }

  return nodemailer.createTransport({
    host: SMTP_HOST,
    port: SMTP_PORT,
    secure: SMTP_SECURE,
    auth: {
      user: SMTP_USER,
      pass: SMTP_PASS,
    },
  });
}

function getSheetsClient() {
  if (!SHEET_ID || !GOOGLE_SERVICE_ACCOUNT_EMAIL || !GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY) {
    return null;
  }

  const auth = new google.auth.JWT({
    email: GOOGLE_SERVICE_ACCOUNT_EMAIL,
    key: GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  return google.sheets({ version: "v4", auth });
}

async function sendSectorEmail({
  contact,
  departmentInput,
  normalizedDept,
  routingCode,
  callerNumber,
  callerName,
  callId,
  source,
  callerType,
  requestObject,
}) {
  const transporter = getTransporter();

  if (!transporter) {
    console.log("EMAIL NOT SENT: SMTP non configuré.");
    return;
  }

  const subject = `Nouvel appel - ${contact.name}`;

  const lines = [
    `Bonjour ${contact.name},`,
    ``,
    `Un appel client a été dirigé vers votre ligne.`,
    ``,
    `Source API : ${source || "non définie"}`,
    `Type appelant : ${callerType || "Non renseigné"}`,
    `Objet : ${requestObject || "Non renseigné"}`,
    `Saisie département : ${departmentInput || "Non renseignée"}`,
    `Département normalisé : ${normalizedDept || "Non renseigné"}`,
    `Code de routage : ${routingCode || "Non renseigné"}`,
    `Numéro appelant : ${callerNumber || "Non remonté"}`,
    `Nom appelant : ${callerName || "Non remonté"}`,
    `ID appel : ${callId || "Non remonté"}`,
    `Numéro routé : ${contact.targetValue}`,
    `Date : ${new Date().toLocaleString("fr-FR", { timeZone: "Europe/Paris" })}`,
    ``,
    `Email automatique généré par l'API de routage Aircall.`,
  ];

  await transporter.sendMail({
    from: MAIL_FROM,
    to: contact.email,
    subject,
    text: lines.join("\n"),
  });

  console.log(`EMAIL SENT TO ${contact.email}`);
}

async function appendRoutingLogToSheet({
  callerNumber,
  departmentCode,
  reason,
  selected,
  selectedEmail,
  targetValue,
  callerType,
  requestObject,
  sourceAgent,
  callId,
  callUuid,
  noteRaw,
}) {
  const sheets = getSheetsClient();

  if (!sheets) {
    console.log("SHEETS NOT WRITTEN: configuration Google manquante.");
    return;
  }

  const values = [[
    new Date().toLocaleString("fr-FR", { timeZone: "Europe/Paris" }), // A
    callerNumber || "", // B
    departmentCode || "", // C
    reason || "", // D
    selected || "", // E
    selectedEmail || "", // F
    targetValue || "", // G
    "en_cours", // H
    0, // I
    callerType || "", // J
    requestObject || "", // K
    sourceAgent || "", // L
    callId || "", // M
    callUuid || "", // N
    noteRaw || "", // O
  ]];

  const response = await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!A:O`,
    valueInputOption: "RAW",
    insertDataOption: "INSERT_ROWS",
    requestBody: { values },
  });

  console.log("GOOGLE SHEETS APPEND OK:", response.status);
  console.log("GOOGLE SHEETS UPDATED RANGE:", response.data?.updates?.updatedRange);
}

async function postRoutingSideEffects({
  result,
  callerNumber,
  callerName,
  source,
  voiceData,
}) {
  await Promise.allSettled([
    sendSectorEmail({
      contact: result.contact,
      departmentInput: result.departmentInput,
      normalizedDept: result.normalizedDept,
      routingCode: result.routingCode,
      callerNumber,
      callerName,
      callId: voiceData.callId,
      source,
      callerType: voiceData.callerType,
      requestObject: voiceData.requestObject,
    }),
    appendRoutingLogToSheet({
      callerNumber,
      departmentCode: result.routingCode || result.normalizedDept,
      reason: result.reason,
      selected: result.contact.name,
      selectedEmail: result.contact.email,
      targetValue: result.contact.targetValue,
      callerType: voiceData.callerType,
      requestObject: voiceData.requestObject,
      sourceAgent: voiceData.sourceAgent,
      callId: voiceData.callId,
      callUuid: voiceData.callUuid,
      noteRaw: voiceData.noteRaw,
    }),
  ]);
}

app.post("/aircall/voice-standard-department", checkAuth, async (req, res) => {
  console.log("=== VOICE STANDARD DEPARTMENT ===");
  console.log(JSON.stringify(req.body, null, 2));

  const rawDepartmentInput =
    req.body.departmentInput ??
    req.body.departmentCode ??
    req.body.postalCode ??
    req.body.dept ??
    req.body.answer ??
    "";

  const rawAttempts = req.body.attempts ?? 0;
  const callerNumber = req.body.callerNumber ?? req.body.from ?? "";
  const callerName =
    req.body.callerName ??
    req.body.name ??
    req.body.caller_name ??
    "";

  const voiceData = extractVoiceAgentData(req.body);
  const result = resolveTarget(rawDepartmentInput, rawAttempts, voiceData.callerType);

  console.log("departmentInput =", rawDepartmentInput);
  console.log("callerType =", voiceData.callerType);
  console.log("normalizedDept =", result.normalizedDept);
  console.log("routingCode =", result.routingCode);
  console.log("selected =", result.contact.name);
  console.log("requestObject =", voiceData.requestObject);

  res.json({
    routing: {
      targetType: result.contact.targetType,
      targetValue: result.contact.targetValue,
    },
    debug: {
      callerType: voiceData.callerType,
      normalizedDept: result.normalizedDept,
      routingCode: result.routingCode,
      selected: result.contact.name,
      reason: result.reason,
      requestObject: voiceData.requestObject,
    },
  });

  postRoutingSideEffects({
    result,
    callerNumber,
    callerName,
    source: "voice-standard-department",
    voiceData,
  }).catch((e) => console.error("POST ROUTING ERROR:", e));
});

app.get("/health", (req, res) => {
  res.json({ ok: true });
});

app.listen(PORT, "0.0.0.0", () => {
  console.log("API running on port " + PORT);
});