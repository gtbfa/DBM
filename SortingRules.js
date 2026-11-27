// ==========================================
// DETERMINISTISCHE SORTIERREGELN & 01_FINANZEN-LOGIK
// ==========================================

function isFinanceFolderPath(folderPath) {
  if (!folderPath) return false;
  const normalized = folderPath.toString().trim().replace(/\s*\/\s*/g, "/").toLowerCase();
  return normalized === "01_finanzen" || normalized.startsWith("01_finanzen/");
}

function isBankStatementDoc(fileName, reasoning) {
  const combined = `${fileName || ""} ${reasoning || ""}`.toLowerCase();
  return ["kontoauszug", "finanzstatus", "umsatzübersicht", "umsatzauskunft"].some(kw => combined.includes(kw));
}

function hasAbrechnungSignal(normalizedText) {
  if (!normalizedText) return false;
  return /\babrechnung\b/.test(normalizedText);
}

function hasGermanInvoiceSignals(normalizedText) {
  if (!normalizedText) return false;

  const nonInvoiceSignals = [
    "abrechnung",
    "kostenvoranschlag",
    "pflegevertrag",
    "abtretungserklaerung",
    "beratungsbesuch",
    "37 abs. 3",
    "37 abs 3",
    "leistungsnachweis"
  ];
  const hasStrongInvoiceWord = [
    "rechnung",
    "rechnungsnr",
    "rechnungsnummer",
    "rechnungsbetrag",
    "gesamtbetrag",
    "endbetrag",
    "zahlbetrag",
    "quittung",
    "kassenbon",
    "gutschrift",
    "honorarrechnung"
  ].some(kw => normalizedText.includes(kw));

  if (hasStrongInvoiceWord) return true;

  const hasEigenanteilInvoiceContext = normalizedText.includes("eigenanteil") &&
    /\b(brutto|mwst|ust|iban|zahlungsziel|faellig|betrag|eur)\b/.test(normalizedText);
  const looksLikeNonInvoiceForm = nonInvoiceSignals.some(kw => normalizedText.includes(kw));

  return hasEigenanteilInvoiceContext && !looksLikeNonInvoiceForm;
}

function getFinancePromptRules() {
  return `STRIKTE SORTIERREGEL FÜR "01_Finanzen" (UND ALLE UNTERORDNER VON "01_Finanzen"):
- Ordner unter "01_Finanzen" (inklusive aller Unterordner!) sind erlaubt für:
  1. Eindeutige RECHNUNGEN, QUITTUNGEN, KASSENBONS oder GUTSCHRIFTEN (deutschsprachige Belege mit Endbetrag/Brutto/Rechnungsbetrag > 0).
  2. AUSNAHME KONTOAUSZÜGE: Bank-Kontoauszüge, Finanzstatus- oder Umsatzübersichten (z. B. Volksbank Kontoauszug) – diese gehören in den entsprechenden Bank-/Volksbank-Ordner unter "01_Finanzen".
- VERBOTEN für "01_Finanzen" sind Dokumente mit "Abrechnung" (soweit keine explizite Rechnung/Quittung/Gutschrift/Kontoauszug vorliegt) sowie Nicht-Rechnungsdokumente wie: Kostenvoranschläge, Pflegeverträge, Abtretungserklärungen, reine Leistungsnachweise ohne Rechnungsbetrag, Nachweise über Beratungsbesuche (§37 Abs. 3 SGB XI), unverbindliche Angebote oder reine Informationsschreiben.
- Solche Dokumente müssen in einen passenden Ordner AUSSERHALB von "01_Finanzen" (z. B. unter "07_Abrechnung" oder "03_Klienten_Verwaltung") sortiert werden.`;
}

function applyFinanceSortingGuardrail(rawParsed, resolvedFolder, altFolder, reasoningText, docType) {
  const hasAccountingFields = typeof rawParsed.t !== "undefined";
  const isValidInvoiceType = !hasAccountingFields || (rawParsed.t === "Einnahme" || rawParsed.t === "Ausgabe");
  const parsedAmount = Number(rawParsed.v);
  const hasPositiveAmount = typeof rawParsed.v === "undefined" || (!isNaN(parsedAmount) && parsedAmount > 0);
  const isUnambiguousInvoice = (rawParsed.is_rechnung === true) && isValidInvoiceType && hasPositiveAmount;
  const isBankStatement = rawParsed.is_kontoauszug === true;

  let nextFolder = resolvedFolder;
  let nextReason = reasoningText;
  let nextDocType = docType;
  let nextAmount = rawParsed.v;

  if (isBankStatement) {
    nextDocType = "Kontoauszug";
    nextAmount = 0;
  } else if (!isUnambiguousInvoice) {
    if (nextDocType === "Einnahme" || nextDocType === "Ausgabe") {
      nextDocType = "Kein Beleg";
    }

    if (isFinanceFolderPath(nextFolder)) {
      const blockedTarget = nextFolder;
      if (altFolder !== "MANUELLE_PRUEFUNG" && !isFinanceFolderPath(altFolder)) {
        nextFolder = altFolder;
        nextReason = `⚠️ Kein Rechnungsbeleg / Abrechnung (${blockedTarget} gesperrt → umgeleitet): ${nextReason}`;
      } else {
        nextFolder = "MANUELLE_PRUEFUNG";
        nextReason = `⚠️ Kein Rechnungsbeleg / Abrechnung (${blockedTarget} gesperrt): ${nextReason}`;
      }
    }
  }

  return {
    folder_name: nextFolder,
    reasoning: nextReason,
    docType: nextDocType,
    amount: nextAmount,
    isBankStatement: isBankStatement
  };
}

function canMoveToFinanceFolder(targetFolderName, isAiSuggestion, bType, bAmount, originalName, newNameInput, reasoning) {
  const isFinanceTarget = isFinanceFolderPath(targetFolderName);
  const isConfirmedInvoice = (bType === "Einnahme" || bType === "Ausgabe") && bAmount > 0;
  const isBankStatement = bType === "Kontoauszug" || isBankStatementDoc(`${originalName} ${newNameInput}`, reasoning);
  const allowed = !(isAiSuggestion && isFinanceTarget && !isConfirmedInvoice && !isBankStatement);

  return {
    allowed: allowed,
    isFinanceTarget: isFinanceTarget,
    isConfirmedInvoice: isConfirmedInvoice,
    isBankStatement: isBankStatement
  };
}

function partitionFinanceCandidateFolders(scoredItems, financeBudget, fallbackBudget, pickedIdx) {
  const matchedFinance = scoredItems
    .filter(item => isFinanceFolderPath(item.path) && item.score > 0)
    .sort((a, b) => b.score - a.score || a.depth - b.depth)
    .slice(0, financeBudget);

  const fallbackNonFinance = scoredItems
    .filter(item => !isFinanceFolderPath(item.path) && !pickedIdx.has(item.idx))
    .slice(0, fallbackBudget);

  return {
    matchedFinance: matchedFinance,
    fallbackNonFinance: fallbackNonFinance
  };
}

const SORTING_RULES = [
  {
    id: "BANK_STATEMENT_VOLKSBANK",
    priority: 100,
    mode: "HARD_ROUTE",
    requiresFinanceCheck: true,
    docTypeHint: "Kontoauszug",
    categoryHint: "Kontoführung & Bankgebühren",
    when: {
      keywordsAny: [
        "kontoauszug",
        "finanzstatus",
        "umsatzübersicht",
        "umsatzuebersicht",
        "umsatzauskunft",
        "rechnungsabschluss",
        "volksbank",
        "vb 2026"
      ],
      excludeKeywords: [
        "eingangsrechnung",
        "ausgangsrechnung",
        "mahnung"
      ]
    },
    then: {
      targetFolderPatterns: [
        "01_Finanzen / 01_Bank_und_Zahlungsverkehr",
        "01_Finanzen / Volks Bank",
        "01_Finanzen / Volksbank"
      ],
      reason: "Regel: Kontoauszug / Finanzstatus"
    }
  },
  {
    id: "SGB11_37_3_BERATUNGSBESUCH_NACHWEIS_CLIENT",
    priority: 95,
    mode: "HARD_ROUTE",
    requiresFinanceCheck: false,
    docTypeHint: "Kein Beleg",
    categoryHint: "Beratungsbesuch (§37 Abs. 3 SGB XI)",
    when: {
      keywordsAny: [
        "beratungsbesuch",
        "37 abs. 3",
        "37 abs 3",
        "§37 abs. 3",
        "nachweis beratungsbesuch"
      ],
      excludeKeywords: [
        "rechnungsbetrag",
        "rechnungsnummer",
        "zahlungsziel"
      ]
    },
    then: {
      targetFolderPatterns: [
        "07_Abrechnung / Beratungsbesuch nach §37 Absatz 3 SGB XI / {client}",
        "03_Klienten_Verwaltung / {client}"
      ],
      reason: "Regel: §37 Abs. 3 Beratungsbesuch-Nachweis ({client})"
    }
  },
  {
    id: "SGB11_37_3_BERATUNGSBESUCH_NACHWEIS_NEW_CLIENT",
    priority: 92,
    mode: "HARD_ROUTE",
    requiresFinanceCheck: false,
    docTypeHint: "Kein Beleg",
    categoryHint: "Beratungsbesuch (§37 Abs. 3 SGB XI)",
    when: {
      requiresUnknownClient: true,
      keywordsAny: [
        "beratungsbesuch",
        "37 abs. 3",
        "37 abs 3",
        "§37 abs. 3",
        "nachweis beratungsbesuch"
      ],
      excludeKeywords: [
        "rechnungsbetrag",
        "rechnungsnummer",
        "zahlungsziel"
      ]
    },
    then: {
      targetFolderPatterns: [
        "07_Abrechnung / Beratungsbesuch nach §37 Absatz 3 SGB XI",
        "MANUELLE_PRUEFUNG"
      ],
      reason: "Regel: §37 Abs. 3 ohne bekannten Klienten – Ordner/Zuordnung prüfen"
    }
  },
  {
    id: "ABRECHNUNG_NON_FINANCE_ROUTING",
    priority: 89,
    mode: "HARD_ROUTE",
    requiresFinanceCheck: false,
    docTypeHint: "Kein Beleg",
    when: {
      customPredicate: "hasAbrechnungSignal",
      excludeKeywords: [
        "rechnung",
        "rechnungsnr",
        "rechnungsnummer",
        "rechnungsbetrag",
        "quittung",
        "kassenbon",
        "gutschrift",
        "kontoauszug",
        "finanzstatus"
      ]
    },
    then: {
      targetFolderPatterns: [
        "07_Abrechnung / {client} / {period}",
        "07_Abrechnung / {client}",
        "03_Klienten_Verwaltung / {client}",
        "07_Abrechnung",
        "MANUELLE_PRUEFUNG"
      ],
      reason: "Regel: 'Abrechnung' (keine Rechnung) → Ziel außerhalb 01_Finanzen"
    }
  },
  {
    id: "INVOICE_FINANCE_CANDIDATE_BOOST",
    priority: 88,
    mode: "BOOST_CANDIDATES",
    requiresFinanceCheck: false,
    when: {
      customPredicate: "hasGermanInvoiceSignals"
    },
    then: {
      boostFolderPatterns: [
        "01_Finanzen / 02_Ausgangsrechnungen / 03-1_Krankenkassen_SGB_V",
        "01_Finanzen / 02_Ausgangsrechnungen / 03-2_Pflegekassen_SGB_XI",
        "01_Finanzen / 02_Ausgangsrechnungen / 03-3_Sozialhilfeträger_SGB_XII",
        "01_Finanzen / 02_Ausgangsrechnungen / 03-4_Privatzahler_und_Eigenanteile",
        "01_Finanzen / 02_Ausgangsrechnungen / 03-5_Absetzungen_Korrekturen_Gutschriften",
        "01_Finanzen / 03_Eingangsrechnungen / 04-1_Pflege_und_Medizinbedarf_Kleidung",
        "01_Finanzen / 03_Eingangsrechnungen / 04-2_Fahrzeugkosten",
        "01_Finanzen / 03_Eingangsrechnungen / 04-3_IT_Software",
        "01_Finanzen / 03_Eingangsrechnungen / 04-4_Raumkosten_und_Bürobetrieb",
        "01_Finanzen / 03_Eingangsrechnungen / 04-5_Versicherungen_Beiträge_Gebühren",
        "01_Finanzen / 03_Eingangsrechnungen / 04-6_Personal_Fortbildung_und_Beratung",
        "01_Finanzen / 03_Eingangsrechnungen / 04-7_Anlagevermögen_und_Investitionen",
        "01_Finanzen / 03_Eingangsrechnungen / 04-8_Belege"
      ],
      reason: "Regel: Rechnungsbeleg erkannt – 01_Finanzen-Ordner priorisiert"
    }
  },
  {
    id: "CLIENT_BILLING_PERIOD_ROUTING",
    priority: 80,
    mode: "BOOST_CANDIDATES",
    requiresFinanceCheck: false,
    when: {
      requiresKnownClient: true
    },
    then: {
      boostFolderPatterns: [
        "07_Abrechnung / {client}",
        "03_Klienten_Verwaltung / {client}",
        "01_Finanzen / 02_Ausgangsrechnungen / 03-2_Pflegekassen_SGB_XI",
        "01_Finanzen / 02_Ausgangsrechnungen / 03-4_Privatzahler_und_Eigenanteile"
      ],
      reason: "Regel: Klientenzuordnung ({client})"
    }
  },
  {
    id: "NEW_CLIENT_CARE_BILLING_REVIEW",
    priority: 75,
    mode: "HARD_ROUTE",
    requiresFinanceCheck: false,
    when: {
      requiresUnknownClient: true,
      keywordsAny: [
        "pflegegrad",
        "sgb xi",
        "sgb v",
        "sgb xii",
        "leistungsnachweis",
        "pflegevertrag",
        "klient",
        "patient",
        "abrechnung"
      ],
      excludeKeywords: [
        "rechnung",
        "rechnungsnr",
        "rechnungsnummer",
        "rechnungsbetrag",
        "gesamtbetrag",
        "endbetrag",
        "brutto",
        "quittung",
        "kassenbon",
        "gutschrift",
        "kontoauszug",
        "finanzstatus",
        "volksbank",
        "domainfactory",
        "google workspace",
        "bgw",
        "gothaer",
        "sbk",
        "aok",
        "dak"
      ]
    },
    then: {
      targetFolderPatterns: [
        "MANUELLE_PRUEFUNG"
      ],
      reason: "Regel: Möglicher neuer Klient / Nicht-Rechnung – manuelle Prüfung"
    }
  },
  {
    id: "VENDOR_KRANKENKASSE_SBK",
    priority: 70,
    mode: "BOOST_CANDIDATES",
    requiresFinanceCheck: false,
    when: {
      keywordsAny: [
        "sbk",
        "siemens-betriebskrankenkasse",
        "siemens betriebskrankenkasse",
        "pflegekasse"
      ]
    },
    then: {
      boostFolderPatterns: [
        "02_Organisation_und_Zulassungen / 01-4_Betriebsversicherungen / Krankenkasse SBK_2026",
        "01_Finanzen / 03_Eingangsrechnungen / 04-5_Versicherungen_Beiträge_Gebühren",
        "01_Finanzen / 02_Ausgangsrechnungen / 03-1_Krankenkassen_SGB_V",
        "01_Finanzen / 02_Ausgangsrechnungen / 03-2_Pflegekassen_SGB_XI"
      ],
      reason: "Regel: Krankenkasse / SBK"
    }
  },
  {
    id: "VENDOR_IT_HOSTING_WORKSPACE",
    priority: 65,
    mode: "BOOST_CANDIDATES",
    requiresFinanceCheck: false,
    when: {
      keywordsAny: [
        "domainfactory",
        "google workspace",
        "google cloud",
        "ionos",
        "strato"
      ]
    },
    then: {
      boostFolderPatterns: [
        "01_Finanzen / 03_Eingangsrechnungen / 04-3_IT_Software",
        "01_Finanzen / 03_Eingangsrechnungen / 04-8_Belege",
        "02_Organisation_und_Zulassungen"
      ],
      reason: "Regel: IT / Domain / Workspace"
    }
  },
  {
    id: "VENDOR_VERSICHERUNG_BGW_AUTO",
    priority: 65,
    mode: "BOOST_CANDIDATES",
    requiresFinanceCheck: false,
    when: {
      keywordsAny: [
        "bgw",
        "gothaer",
        "berufshaftpflicht",
        "kfz-versicherung",
        "kfz versicherung",
        "tankstelle",
        "werkstatt"
      ]
    },
    then: {
      boostFolderPatterns: [
        "01_Finanzen / 03_Eingangsrechnungen / 04-5_Versicherungen_Beiträge_Gebühren",
        "01_Finanzen / 03_Eingangsrechnungen / 04-2_Fahrzeugkosten",
        "02_Organisation_und_Zulassungen / 01-4_Betriebsversicherungen / Berufshaftpflichtversicherung"
      ],
      reason: "Regel: Versicherung / BGW / Auto"
    }
  }
];

const KNOWN_CLIENTS = [
  { canonical: "Fouad Sebky", variants: ["fouad sebky", "sebky fouad", "fouad_sebky", "sebky_fouad"] },
  { canonical: "Boris Langer", variants: ["boris langer", "langer boris", "boris_langer", "langer_boris"] },
  { canonical: "Dieter Reichel", variants: ["dieter reichel", "reichel dieter", "dieter_reichel", "reichel_dieter"] },
  { canonical: "Eva Rotter", variants: ["eva rotter", "rotter eva", "eva_rotter", "rotter_eva"] },
  { canonical: "Henriette Deligne", variants: ["henriette deligne", "deligne henriette", "henriette_deligne"] },
  { canonical: "Margret Obermeier-Hartmann", variants: ["margret obermeier-hartmann", "margret obermeier hartmann", "obermeier-hartmann", "obermeier hartmann"] },
  { canonical: "Gerd Obermeier-Hartmann", variants: ["gerd obermeier-hartmann", "gerd obermeier hartmann", "gerd_obermeier"] },
  { canonical: "Renate Peters", variants: ["renate peters", "peters renate", "renate_peters"] },
  { canonical: "Yvonne Schwerdtfeger", variants: ["yvonne schwerdtfeger", "schwerdtfeger yvonne", "yvonne_schwerdtfeger"] },
  { canonical: "Hak Abu Haj", variants: ["hak abu haj", "hak-abu-haj", "hak_abu_haj", "abu haj"] },
  { canonical: "Birgit Hösemann", variants: ["birgit hösemann", "birgit hoesemann", "hösemann", "hoesemann"] },
  { canonical: "Hussein Taleb", variants: ["hussein taleb", "taleb hussein", "hussein_taleb"] },
  { canonical: "Zahraa Taleb", variants: ["zahraa taleb", "taleb zahraa", "zahraa_taleb"] }
];

function normalizeRuleText(raw) {
  if (!raw) return "";
  return raw
    .toString()
    .toLowerCase()
    .replace(/ä/g, "ae")
    .replace(/ö/g, "oe")
    .replace(/ü/g, "ue")
    .replace(/ß/g, "ss")
    .replace(/[_/\-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeFolderKey(path) {
  return normalizeRuleText(path).replace(/\s+/g, " ");
}

function extractDynamicClientsFromFolderMap(folderMap) {
  if (!folderMap || typeof folderMap !== "object") return [];
  const dynamicClients = [];
  const seenCanonicals = new Set(KNOWN_CLIENTS.map(c => normalizeRuleText(c.canonical)));
  const clientRootPrefixes = [
    "03_klienten_verwaltung/",
    "07_abrechnung/"
  ];
  const ignoredLeafTokens = new Set([
    "beratungsbesuch nach §37 absatz 3 sgb xi",
    "beratungsbesuch nach 37 absatz 3 sgb xi",
    "pflegekassen sgb xi",
    "privatzahler und eigenanteile",
    "abrechnung",
    "rechnungen",
    "belege"
  ]);

  Object.keys(folderMap).forEach(fullPath => {
    const normPath = normalizeFolderKey(fullPath);
    for (let p = 0; p < clientRootPrefixes.length; p++) {
      const prefix = clientRootPrefixes[p];
      if (!normPath.startsWith(prefix)) continue;

      const remainder = normPath.slice(prefix.length).trim();
      if (!remainder) continue;
      const firstSegment = remainder.split("/")[0].trim();
      if (!firstSegment || firstSegment.length < 4 || ignoredLeafTokens.has(firstSegment)) continue;
      if (/^(januar|februar|maerz|marz|april|mai|juni|juli|august|september|oktober|november|dezember) /.test(firstSegment)) {
        continue;
      }

      if (!seenCanonicals.has(firstSegment)) {
        seenCanonicals.add(firstSegment);
        const rawSegments = fullPath.split("/").map(s => s.trim());
        const rootIdx = rawSegments.findIndex(seg => {
          const n = normalizeRuleText(seg);
          return n === "03 klienten verwaltung" || n === "07 abrechnung";
        });
        const displayName = (rootIdx >= 0 && rawSegments[rootIdx + 1]) ? rawSegments[rootIdx + 1].trim() : firstSegment;
        dynamicClients.push({
          canonical: displayName,
          variants: [firstSegment]
        });
      }
    }
  });

  return dynamicClients;
}

function detectClientName(normalizedHaystack, folderMap, extraCandidateName) {
  const allClients = KNOWN_CLIENTS.concat(extractDynamicClientsFromFolderMap(folderMap));

  if (normalizedHaystack) {
    for (let i = 0; i < allClients.length; i++) {
      const client = allClients[i];
      for (let v = 0; v < client.variants.length; v++) {
        const normVar = normalizeRuleText(client.variants[v]);
        if (normVar && normalizedHaystack.includes(normVar)) {
          return { name: client.canonical, isKnown: true };
        }
      }
    }
  }

  const cleanedExtra = normalizeRuleText(extraCandidateName);
  if (cleanedExtra && cleanedExtra.length >= 3) {
    for (let i = 0; i < allClients.length; i++) {
      const client = allClients[i];
      const normCanon = normalizeRuleText(client.canonical);
      if (normCanon === cleanedExtra || client.variants.some(v => normalizeRuleText(v) === cleanedExtra)) {
        return { name: client.canonical, isKnown: true };
      }
    }
    return { name: extraCandidateName.toString().trim(), isKnown: false };
  }

  return { name: "", isKnown: false };
}

function detectBillingPeriodLabel(normalizedHaystack) {
  if (!normalizedHaystack) return "";
  const monthNames = [
    ["januar", "Januar"], ["februar", "Februar"], ["maerz", "März"], ["marz", "März"],
    ["april", "April"], ["mai", "Mai"], ["juni", "Juni"],
    ["juli", "Juli"], ["august", "August"], ["september", "September"],
    ["oktober", "Oktober"], ["november", "November"], ["dezember", "Dezember"]
  ];

  for (let i = 0; i < monthNames.length; i++) {
    const [key, label] = monthNames[i];
    const re = new RegExp(`\\b${key}(?:\\s*202[4-9]|\\s*2[4-9])?\\b`, "i");
    if (re.test(normalizedHaystack)) {
      return label;
    }
  }

  const mmYYYY = normalizedHaystack.match(/\b(0[1-9]|1[0-2])[.\/\s-](202[4-9])\b/);
  if (mmYYYY) {
    const idx = parseInt(mmYYYY[1], 10) - 1;
    return monthNames[idx] ? monthNames[idx][1] : mmYYYY[0];
  }
  return "";
}

function buildSortingContext(file, extractedText, subFolderPath, folderMap, extraCandidateName) {
  const fileName = file && typeof file.getName === "function" ? file.getName() : "";
  const rawCombined = [subFolderPath || "", fileName, extractedText || ""].join(" ");
  const normalizedAll = normalizeRuleText(rawCombined);
  const normalizedHeader = normalizeRuleText([subFolderPath || "", fileName].join(" "));
  const clientInfo = detectClientName(normalizedAll, folderMap, extraCandidateName);

  return {
    fileName: fileName,
    subFolderPath: subFolderPath || "",
    extractedText: extractedText || "",
    normalizedAll: normalizedAll,
    normalizedHeader: normalizedHeader,
    clientName: clientInfo.name,
    isKnownClient: clientInfo.isKnown,
    periodLabel: detectBillingPeriodLabel(normalizedAll),
    isInvoiceCandidate: hasGermanInvoiceSignals(normalizedAll),
    isAbrechnungDoc: hasAbrechnungSignal(normalizedAll)
  };
}

function resolveFolderPattern(pattern, ctx, folderMap) {
  if (!pattern) return "";
  const trimmedPattern = pattern.trim();
  if (trimmedPattern === "MANUELLE_PRUEFUNG") {
    return "MANUELLE_PRUEFUNG";
  }
  if (!folderMap) return "";

  const populated = trimmedPattern
    .replace(/\{client\}/g, ctx.clientName || "")
    .replace(/\{period\}/g, ctx.periodLabel || "")
    .replace(/\s+\/\s+$/g, "")
    .replace(/\s+/g, " ")
    .trim();

  if (populated.includes("{client}") || populated.includes("{period}")) {
    return "";
  }

  if (folderMap[populated]) return populated;

  const targetNorm = normalizeFolderKey(populated);
  const keys = Object.keys(folderMap);
  for (let i = 0; i < keys.length; i++) {
    if (normalizeFolderKey(keys[i]) === targetNorm) {
      return keys[i];
    }
  }
  return "";
}

function matchesRuleConditions(whenCfg, ctx) {
  if (!whenCfg) return false;
  const haystack = ctx.normalizedAll;
  if (!haystack) return false;

  if (whenCfg.requiresKnownClient === true && !ctx.isKnownClient) {
    return false;
  }

  if (whenCfg.requiresUnknownClient === true && ctx.isKnownClient) {
    return false;
  }

  if (whenCfg.customPredicate === "hasGermanInvoiceSignals" && !ctx.isInvoiceCandidate) {
    return false;
  }

  if (whenCfg.customPredicate === "hasAbrechnungSignal" && !ctx.isAbrechnungDoc) {
    return false;
  }

  if (Array.isArray(whenCfg.excludeKeywords)) {
    for (let i = 0; i < whenCfg.excludeKeywords.length; i++) {
      const ex = normalizeRuleText(whenCfg.excludeKeywords[i]);
      if (ex && haystack.includes(ex)) return false;
    }
  }

  if (Array.isArray(whenCfg.keywordsAll) && whenCfg.keywordsAll.length > 0) {
    for (let i = 0; i < whenCfg.keywordsAll.length; i++) {
      const reqKw = normalizeRuleText(whenCfg.keywordsAll[i]);
      if (reqKw && !haystack.includes(reqKw)) return false;
    }
  }

  if (Array.isArray(whenCfg.keywordsAny) && whenCfg.keywordsAny.length > 0) {
    let anyHit = false;
    for (let i = 0; i < whenCfg.keywordsAny.length; i++) {
      const kw = normalizeRuleText(whenCfg.keywordsAny[i]);
      if (kw && haystack.includes(kw)) {
        anyHit = true;
        break;
      }
    }
    if (!anyHit) return false;
  }

  return true;
}

function evaluateSortingRules(ctx, folderMap) {
  const sortedRules = SORTING_RULES.slice().sort((a, b) => (b.priority || 0) - (a.priority || 0));
  const boostedFolders = [];
  const hints = [];

  for (let i = 0; i < sortedRules.length; i++) {
    const rule = sortedRules[i];
    if (!matchesRuleConditions(rule.when, ctx)) continue;

    const reasonText = (rule.then && rule.then.reason ? rule.then.reason : `Regel: ${rule.id}`)
      .replace(/\{client\}/g, ctx.clientName || "Unbekannt")
      .replace(/\{period\}/g, ctx.periodLabel || "");

    if (rule.mode === "HARD_ROUTE" && rule.then && Array.isArray(rule.then.targetFolderPatterns)) {
      for (let p = 0; p < rule.then.targetFolderPatterns.length; p++) {
        const resolvedFolder = resolveFolderPattern(rule.then.targetFolderPatterns[p], ctx, folderMap);
        if (!resolvedFolder) continue;

        const isFinance = isFinanceFolderPath(resolvedFolder);
        if (rule.requiresFinanceCheck && isFinance && rule.docTypeHint !== "Kontoauszug") {
          continue;
        }
        if (rule.id === "ABRECHNUNG_NON_FINANCE_ROUTING" && isFinance) {
          continue;
        }

        const isBankStmt = rule.docTypeHint === "Kontoauszug";
        return {
          matched: true,
          mode: "HARD_ROUTE",
          ruleId: rule.id,
          folder_name: resolvedFolder,
          reasoning: `${reasonText} [0 Tokens]`,
          tokens: 0,
          tokenDisplay: "0 (Regel)",
          t: rule.docTypeHint || "",
          k: rule.categoryHint || "",
          p: ctx.clientName || "",
          v: isBankStmt ? 0 : undefined,
          is_kontoauszug: isBankStmt,
          boostedFolders: [],
          promptHint: ""
        };
      }
    }

    if (rule.mode === "BOOST_CANDIDATES" && rule.then && Array.isArray(rule.then.boostFolderPatterns)) {
      for (let b = 0; b < rule.then.boostFolderPatterns.length; b++) {
        const resolvedBoost = resolveFolderPattern(rule.then.boostFolderPatterns[b], ctx, folderMap);
        if (resolvedBoost && !boostedFolders.includes(resolvedBoost)) {
          boostedFolders.push(resolvedBoost);
        }
      }
      hints.push(reasonText);
    }
  }

  return {
    matched: boostedFolders.length > 0 || hints.length > 0,
    mode: boostedFolders.length > 0 ? "BOOST_CANDIDATES" : "FALLBACK_LLM",
    ruleId: hints.length > 0 ? hints.join("; ") : "",
    folder_name: "",
    reasoning: "",
    tokens: 0,
    tokenDisplay: "0",
    t: "",
    k: "",
    p: ctx.clientName || "",
    boostedFolders: boostedFolders,
    promptHint: hints.length > 0
      ? `Regel-Hinweise: ${hints.join(" | ")}${ctx.periodLabel ? ` | Zeitraum: ${ctx.periodLabel}` : ""}`
      : ""
  };
}