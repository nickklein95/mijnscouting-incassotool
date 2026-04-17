const BANK_TO_BIC = (typeof BIC_LIST !== "undefined" && BIC_LIST) || {};

const REQUIRED_FIELDS = {
  creditorIban: "IBAN Scouting Groep",
  creditorBic: "BIC Scouting Groep",
  creditorName: "Houdernaam Scouting Groep",
  debtorIban: "IBAN Debiteur",
  debtorBic: "BIC Debiteur",
  debtorName: "Naam debiteur",
  amount: "Openstaand bedrag",
  description: "Beschrijving bankafschrift",
  mandateId: "Mandaat Id",
  mandateDate: "Datum mandaat getekend",
  sequenceType: "Type reeks",
  dueDate: "Vervaldatum",
  invoiceNumber: "Factuurnummer",
  paymentMethod: "Betaal methode",
  status: "Status",
  creditorSchemeId: "IncassantenId",
};

const FIELD_LIMITS = {
  messageId: 35,
  paymentInfoId: 35,
  name: 70,
  endToEndId: 35,
  mandateId: 35,
  remittance: 140,
  iban: 34,
  bic: 11,
  creditorSchemeId: 35,
};

const els = {
  inputCard: document.getElementById("inputCard"),
  previewCard: document.getElementById("previewCard"),
  previewTableWrap: document.getElementById("previewTableWrap"),
  fileInput: document.getElementById("fileInput"),
  uploadTrigger: document.getElementById("uploadTrigger"),
  selectedFileName: document.getElementById("selectedFileName"),
  collectionDate: document.getElementById("collectionDate"),
  fileNamePrefix: document.getElementById("fileNamePrefix"),
  generateBtn: document.getElementById("generateBtn"),
  downloadBtn: document.getElementById("downloadBtn"),
  downloadLink: document.getElementById("downloadLink"),
  previewBody: document.getElementById("previewBody"),
  status: document.getElementById("status"),
  txCount: document.getElementById("txCount"),
  ctrlSum: document.getElementById("ctrlSum"),
};

const state = {
  rows: [],
  transactions: [],
  creditor: null,
  xml: "",
  fileName: "",
};

const resizeObserver = new ResizeObserver(() => {
  syncPreviewHeight();
});

resizeObserver.observe(els.inputCard);
window.addEventListener("resize", syncPreviewHeight);
requestAnimationFrame(syncPreviewHeight);

els.fileInput.addEventListener("change", async (event) => {
  const [file] = event.target.files;
  if (!file) {
    resetState();
    event.target.value = "";
    return;
  }
  await handleSelectedFile(file);
  event.target.value = "";
});

for (const eventName of ["dragenter", "dragover"]) {
  els.uploadTrigger.addEventListener(eventName, (event) => {
    event.preventDefault();
    els.uploadTrigger.classList.add("is-dragover");
  });
}

for (const eventName of ["dragleave", "dragend", "drop"]) {
  els.uploadTrigger.addEventListener(eventName, (event) => {
    event.preventDefault();
    if (eventName === "dragleave" && els.uploadTrigger.contains(event.relatedTarget)) {
      return;
    }
    els.uploadTrigger.classList.remove("is-dragover");
  });
}

els.uploadTrigger.addEventListener("drop", async (event) => {
  const [file] = event.dataTransfer?.files || [];
  if (!file) {
    return;
  }
  await handleSelectedFile(file);
});

els.generateBtn.addEventListener("click", () => {
  try {
    const createdAt = new Date();
    const xml = buildXml(createdAt);
    const batchDate = formatIsoDate(createdAt);
    const batchTime = formatTimeForFileName(createdAt);
    const prefix = (els.fileNamePrefix.value || "SEPA-DD").trim() || "SEPA-DD";
    state.xml = xml;
    state.fileName = `${prefix}${batchDate}-${batchTime}.xml`;
    const blob = new Blob([xml], { type: "application/xml;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    els.downloadLink.href = url;
    els.downloadLink.download = state.fileName;
    els.downloadLink.hidden = false;
    els.downloadLink.textContent = "";
    els.downloadBtn.disabled = false;
    setStatus(
      `XML succesvol gegenereerd.\nBestandsnaam: ${state.fileName}`,
      "ok",
    );
  } catch (error) {
    setStatus(error.message || "Het genereren van XML is mislukt.", "error");
  }
});

els.downloadBtn.addEventListener("click", () => {
  if (!state.xml || !els.downloadLink.href) {
    return;
  }
  els.downloadLink.click();
});

function hydrateFromRows(rows) {
  if (!rows.length) {
    throw new Error("Het werkboek bevat geen gegevensregels.");
  }

  validateHeaders(rows[0]);
  const candidateRows = rows
    .map((row, index) => ({ row, rowNumber: index + 2 }))
    .filter(({ row }) => shouldParseRow(row));

  state.rows = candidateRows.map(({ row }) => row);
  const normalisedRows = candidateRows.map(({ row, rowNumber }) =>
    normaliseRow(row, rowNumber),
  );
  const invalidRows = normalisedRows.filter((tx) => tx.error);
  if (invalidRows.length) {
    throw new Error(
      invalidRows.map((tx) => `Row ${tx.rowNumber}: ${tx.error}`).join("\n"),
    );
  }

  state.transactions = normalisedRows.filter((tx) => tx.include);

  if (!state.transactions.length) {
    throw new Error("Er zijn na het filteren geen incassoregels gevonden.");
  }

  state.creditor = {
    iban: state.transactions[0].creditorIban,
    bic: state.transactions[0].creditorBic,
    name: state.transactions[0].creditorName,
    schemeId: state.transactions[0].creditorSchemeId,
  };

  if (!els.collectionDate.value) {
    els.collectionDate.value = state.transactions[0].dueDate;
  }

  renderPreview();
  els.generateBtn.disabled = false;
  const warnings = [];
  const bicFallbackCount = state.transactions.filter(
    (tx) => tx.debtorBicSource === "fallback",
  ).length;
  if (bicFallbackCount > 0) {
    warnings.push(
      `${bicFallbackCount} debiteur-BIC-waarde(n) zijn niet gevonden en krijgen NOTPROVIDED.`,
    );
  }
  setStatus(
    warnings.length
      ? warnings.join("\n")
      : `${state.transactions.length} transacties geladen uit het werkboek.`,
    warnings.length ? "warn" : "ok",
  );
  syncPreviewHeight();
}

function resetState() {
  state.rows = [];
  state.transactions = [];
  state.creditor = null;
  state.xml = "";
  state.fileName = "";
  els.generateBtn.disabled = true;
  els.downloadBtn.disabled = true;
  els.previewBody.innerHTML =
    '<tr><td colspan="4">Geen bestand geladen.</td></tr>';
  els.txCount.textContent = "0";
  els.ctrlSum.textContent = "EUR 0.00";
  setSelectedFileName("");
  resetDownload();
  setStatus("Selecteer een Excel-bestand om te beginnen.", "ok");
  syncPreviewHeight();
}

function resetDownload() {
  if (els.downloadLink.href) {
    URL.revokeObjectURL(els.downloadLink.href);
  }
  els.downloadLink.hidden = true;
  els.downloadLink.textContent = "";
  els.downloadLink.removeAttribute("href");
  els.downloadLink.removeAttribute("download");
  els.downloadBtn.disabled = true;
}

function validateHeaders(row) {
  const missing = Object.values(REQUIRED_FIELDS).filter(
    (header) => !(header in row),
  );
  if (missing.length) {
    throw new Error(`Ontbrekende verwachte kolom(men): ${missing.join(", ")}`);
  }
}

function normaliseRow(row, rowNumber) {
  try {
    const paymentMethod = String(row[REQUIRED_FIELDS.paymentMethod] || "")
      .trim()
      .toLowerCase();
    if (!paymentMethod) {
      throw new Error(`"${REQUIRED_FIELDS.paymentMethod}" ontbreekt`);
    }
    if (paymentMethod !== "incasso") {
      return {
        include: false,
        rowNumber,
      };
    }

    const amount = parseEuroAmount(
      row[REQUIRED_FIELDS.amount],
      REQUIRED_FIELDS.amount,
    );
    const debtorIban = requireMaxLength(
      cleanIban(row[REQUIRED_FIELDS.debtorIban]),
      REQUIRED_FIELDS.debtorIban,
      FIELD_LIMITS.iban,
    );
    const debtorName = requireMaxLength(
      cleanText(row[REQUIRED_FIELDS.debtorName]),
      REQUIRED_FIELDS.debtorName,
      FIELD_LIMITS.name,
    );
    const mandateId = normaliseMandateId(
      row[REQUIRED_FIELDS.mandateId],
      REQUIRED_FIELDS.mandateId,
    );
    const invoiceNumber = requireMaxLength(
      cleanText(row[REQUIRED_FIELDS.invoiceNumber]),
      REQUIRED_FIELDS.invoiceNumber,
      FIELD_LIMITS.endToEndId,
    );
    const dueDate = parseFlexibleDate(
      row[REQUIRED_FIELDS.dueDate],
      REQUIRED_FIELDS.dueDate,
    );

    if (!debtorIban) {
      throw new Error(`"${REQUIRED_FIELDS.debtorIban}" ontbreekt`);
    }
    if (!debtorName) {
      throw new Error(`"${REQUIRED_FIELDS.debtorName}" ontbreekt`);
    }
    if (!mandateId) {
      throw new Error(`"${REQUIRED_FIELDS.mandateId}" ontbreekt`);
    }
    if (!invoiceNumber) {
      throw new Error(`"${REQUIRED_FIELDS.invoiceNumber}" ontbreekt`);
    }
    if (amount <= 0) {
      throw new Error(`"${REQUIRED_FIELDS.amount}" moet groter zijn dan 0`);
    }

    const debtorBicValue = cleanText(
      row[REQUIRED_FIELDS.debtorBic],
    ).toUpperCase();
    const derivedBic = deriveBicFromIban(debtorIban);

    let debtorBic = "";
    let debtorBicSource = "missing";
    if (debtorBicValue && debtorBicValue !== "NONE") {
      debtorBic = requireBicLength(debtorBicValue, REQUIRED_FIELDS.debtorBic);
      debtorBicSource = "sheet";
    } else if (derivedBic) {
      debtorBic = requireBicLength(
        derivedBic,
        `${REQUIRED_FIELDS.debtorBic} (afgeleid)`,
      );
      debtorBicSource = "iban";
    } else {
      debtorBic = "NOTPROVIDED";
      debtorBicSource = "fallback";
    }

    return {
      include: true,
      rowNumber,
      creditorIban: requireMaxLength(
        cleanIban(
          requireText(
            row[REQUIRED_FIELDS.creditorIban],
            REQUIRED_FIELDS.creditorIban,
          ),
        ),
        REQUIRED_FIELDS.creditorIban,
        FIELD_LIMITS.iban,
      ),
      creditorBic: normaliseCreditorBic(row),
      creditorName: requireMaxLength(
        requireText(
          row[REQUIRED_FIELDS.creditorName],
          REQUIRED_FIELDS.creditorName,
        ),
        REQUIRED_FIELDS.creditorName,
        FIELD_LIMITS.name,
      ),
      creditorSchemeId: requireMaxLength(
        requireText(
          row[REQUIRED_FIELDS.creditorSchemeId],
          REQUIRED_FIELDS.creditorSchemeId,
        ).toUpperCase(),
        REQUIRED_FIELDS.creditorSchemeId,
        FIELD_LIMITS.creditorSchemeId,
      ),
      debtorIban,
      debtorBic,
      debtorBicSource,
      debtorName,
      amount,
      description: enforceMaxLength(
        cleanText(row[REQUIRED_FIELDS.description]),
        REQUIRED_FIELDS.description,
        FIELD_LIMITS.remittance,
      ),
      mandateId,
      mandateDate: parseFlexibleDate(
        row[REQUIRED_FIELDS.mandateDate],
        REQUIRED_FIELDS.mandateDate,
      ),
      dueDate,
      invoiceNumber,
    };
  } catch (error) {
    return {
      include: false,
      rowNumber,
      error: error.message || "Ongeldige gegevens",
    };
  }
}

function shouldParseRow(row) {
  const paymentMethod = cleanText(
    row[REQUIRED_FIELDS.paymentMethod],
  ).toLowerCase();
  const debtorName = cleanText(row[REQUIRED_FIELDS.debtorName]);
  const debtorIban = cleanIban(row[REQUIRED_FIELDS.debtorIban]);
  const mandateId = cleanText(row[REQUIRED_FIELDS.mandateId]);
  const invoiceNumber = cleanText(row[REQUIRED_FIELDS.invoiceNumber]);
  const amountRaw = cleanText(row[REQUIRED_FIELDS.amount]);

  if (paymentMethod === "incasso") {
    return true;
  }

  const signals = [
    debtorName,
    debtorIban,
    mandateId,
    invoiceNumber,
    amountRaw,
  ].filter(Boolean).length;

  return signals >= 3;
}

function renderPreview() {
  els.txCount.textContent = String(state.transactions.length);
  els.ctrlSum.textContent = `EUR ${formatAmount(sumAmounts(state.transactions))}`;

  const previewRows = state.transactions
    .map(
      (tx) => `
    <tr>
      <td>${escapeHtml(tx.debtorName)}</td>
      <td>${escapeHtml(tx.debtorIban)}</td>
      <td>${escapeHtml(formatAmount(tx.amount))}</td>
      <td>${escapeHtml(tx.description)}</td>
    </tr>
  `,
    )
    .join("");

  els.previewBody.innerHTML = previewRows;
  syncPreviewHeight();
}

function buildXml(createdAt = new Date()) {
  if (!state.transactions.length || !state.creditor) {
    throw new Error("Laad een werkboek voordat je XML genereert.");
  }

  const collectionDate = validateDateInput(
    els.collectionDate.value,
    "Incassodatum",
  );
  const localInstrument = "CORE";
  const creditorName = requireMaxLength(
    state.creditor.name,
    "Naam crediteur",
    FIELD_LIMITS.name,
  );
  const initiatingParty = requireMaxLength(
    state.creditor.name,
    "Naam initiërende partij",
    FIELD_LIMITS.name,
  );
  const messageId = requireMaxLength(
    formatCompactMessageId(createdAt),
    "MsgId",
    FIELD_LIMITS.messageId,
  );

  const transactions = state.transactions.map((tx) => ({
    ...tx,
    creditorName,
    creditorBic: state.creditor.bic,
    localInstrument,
  }));

  const ctrlSum = formatAmount(sumAmounts(transactions));
  const txCount = String(transactions.length);

  const pmtInfXml = buildPaymentInfoXml({
    messageId,
    collectionDate,
    creditorName,
    localInstrument,
    transactions,
  });

  return `<?xml version="1.0" encoding="UTF-8" ?>\n<Document xmlns="urn:iso:std:iso:20022:tech:xsd:pain.008.001.02" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n  <CstmrDrctDbtInitn>\n    <GrpHdr>\n      <MsgId>${xml(messageId)}</MsgId>\n      <CreDtTm>${xml(formatIsoDateTime(createdAt))}</CreDtTm>\n      <NbOfTxs>${xml(txCount)}</NbOfTxs>\n      <CtrlSum>${xml(ctrlSum)}</CtrlSum>\n      <InitgPty>\n        <Nm>${xml(initiatingParty)}</Nm>\n      </InitgPty>\n    </GrpHdr>\n${pmtInfXml}\n  </CstmrDrctDbtInitn>\n</Document>\n`;
}

function buildPaymentInfoXml({
  messageId,
  collectionDate,
  creditorName,
  localInstrument,
  transactions,
}) {
  const ctrlSum = formatAmount(sumAmounts(transactions));
  const pmtInfId = requireMaxLength(
    `BATCH-${collectionDate}`,
    "PmtInfId",
    FIELD_LIMITS.paymentInfoId,
  );
  const txXml = transactions
    .map((tx) => {
      const dbtrAgt = tx.debtorBic
        ? `        <DbtrAgt>\n          <FinInstnId>\n${tx.debtorBic === "NOTPROVIDED" ? "            <Othr><Id>NOTPROVIDED</Id></Othr>" : `            <BIC>${xml(tx.debtorBic)}</BIC>`}\n          </FinInstnId>\n        </DbtrAgt>\n`
        : "";

      return `      <DrctDbtTxInf>\n        <PmtId>\n          <EndToEndId>${xml(tx.invoiceNumber)}</EndToEndId>\n        </PmtId>\n        <InstdAmt Ccy="EUR">${xml(formatAmount(tx.amount))}</InstdAmt>\n        <DrctDbtTx>\n          <MndtRltdInf>\n            <MndtId>${xml(tx.mandateId)}</MndtId>\n            <DtOfSgntr>${xml(tx.mandateDate)}</DtOfSgntr>\n          </MndtRltdInf>\n        </DrctDbtTx>\n${dbtrAgt}        <Dbtr>\n          <Nm>${xml(tx.debtorName)}</Nm>\n        </Dbtr>\n        <DbtrAcct>\n          <Id>\n            <IBAN>${xml(tx.debtorIban)}</IBAN>\n          </Id>\n        </DbtrAcct>\n        <RmtInf>\n          <Ustrd>${xml(tx.description || tx.invoiceNumber)}</Ustrd>\n        </RmtInf>\n      </DrctDbtTxInf>`;
    })
    .join("\n");

  const creditorBicXml = state.creditor.bic
    ? `      <CdtrAgt>\n        <FinInstnId>\n          <BIC>${xml(state.creditor.bic)}</BIC>\n        </FinInstnId>\n      </CdtrAgt>\n`
    : "";

  return `    <PmtInf>\n      <PmtInfId>${xml(pmtInfId)}</PmtInfId>\n      <PmtMtd>DD</PmtMtd>\n      <NbOfTxs>${xml(String(transactions.length))}</NbOfTxs>\n      <CtrlSum>${xml(ctrlSum)}</CtrlSum>\n      <PmtTpInf>\n        <SvcLvl>\n          <Cd>SEPA</Cd>\n        </SvcLvl>\n        <LclInstrm>\n          <Cd>${xml(localInstrument)}</Cd>\n        </LclInstrm>\n        <SeqTp>RCUR</SeqTp>\n      </PmtTpInf>\n      <ReqdColltnDt>${xml(collectionDate)}</ReqdColltnDt>\n      <Cdtr>\n        <Nm>${xml(creditorName)}</Nm>\n      </Cdtr>\n      <CdtrAcct>\n        <Id>\n          <IBAN>${xml(state.creditor.iban)}</IBAN>\n        </Id>\n      </CdtrAcct>\n${creditorBicXml}      <ChrgBr>SLEV</ChrgBr>\n      <CdtrSchmeId>\n        <Id>\n          <PrvtId>\n            <Othr>\n              <Id>${xml(state.creditor.schemeId)}</Id>\n              <SchmeNm>\n                <Prtry>SEPA</Prtry>\n              </SchmeNm>\n            </Othr>\n          </PrvtId>\n        </Id>\n      </CdtrSchmeId>\n${txXml}\n    </PmtInf>`;
}

function deriveBicFromIban(iban) {
  if (!iban || iban.length < 8 || !iban.startsWith("NL")) {
    return "";
  }
  const bankCode = iban.slice(4, 8).toUpperCase();
  return BANK_TO_BIC[bankCode] || "";
}

function normaliseCreditorBic(row) {
  const creditorIban = cleanIban(row[REQUIRED_FIELDS.creditorIban]);
  const creditorBicValue = cleanText(
    row[REQUIRED_FIELDS.creditorBic],
  ).toUpperCase();
  const creditorBic = creditorBicValue || deriveBicFromIban(creditorIban) || "";
  return creditorBic
    ? requireBicLength(creditorBic, REQUIRED_FIELDS.creditorBic)
    : "";
}

function normaliseMandateId(value, label) {
  const text = requireText(value, label)
    .replace(/_/g, "-")
    .replace(/\s+/g, " ")
    .trim();
  return text.slice(0, FIELD_LIMITS.mandateId);
}

function parseEuroAmount(value, label = "Bedrag") {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }
  const text = String(value || "").trim();
  if (!text) {
    throw new Error(`"${label}" ontbreekt`);
  }
  const stripped = text
    .replace(/\s+/g, "")
    .replace(/^EUR/i, "")
    .replace(/[^\d,.-]/g, "");
  const hasComma = stripped.includes(",");
  const hasDot = stripped.includes(".");
  let normalised = stripped;

  if (hasComma && hasDot) {
    if (stripped.lastIndexOf(",") > stripped.lastIndexOf(".")) {
      normalised = stripped.replace(/\./g, "").replace(",", ".");
    } else {
      normalised = stripped.replace(/,/g, "");
    }
  } else if (hasComma) {
    normalised = stripped.replace(",", ".");
  }

  const amount = Number(normalised);
  if (!Number.isFinite(amount)) {
    throw new Error(`Ongeldige waarde voor "${label}": ${text}`);
  }
  return Math.round(amount * 100) / 100;
}

function parseFlexibleDate(value, label = "Datum") {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return formatIsoDate(value);
  }
  const raw = String(value || "").trim();
  if (!raw) {
    throw new Error(`"${label}" ontbreekt`);
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    return raw;
  }
  const ddmmyyyy = raw.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (ddmmyyyy) {
    return `${ddmmyyyy[3]}-${ddmmyyyy[2]}-${ddmmyyyy[1]}`;
  }
  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) {
    return formatIsoDate(parsed);
  }
  throw new Error(`Ongeldige waarde voor "${label}": ${raw}`);
}

function validateDateInput(value, label) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value || "")) {
    throw new Error(`${label} is verplicht.`);
  }
  return value;
}

function formatIsoDate(date) {
  return [
    date.getFullYear(),
    String(date.getMonth() + 1).padStart(2, "0"),
    String(date.getDate()).padStart(2, "0"),
  ].join("-");
}

function formatIsoDateTime(date) {
  return `${formatIsoDate(date)}T${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")}:${String(date.getSeconds()).padStart(2, "0")}`;
}

function formatTimeForFileName(date) {
  return `${String(date.getHours()).padStart(2, "0")}-${String(date.getMinutes()).padStart(2, "0")}`;
}

function formatCompactTimestamp(date) {
  return `${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, "0")}${String(date.getDate()).padStart(2, "0")}-${String(date.getHours()).padStart(2, "0")}${String(date.getMinutes()).padStart(2, "0")}${String(date.getSeconds()).padStart(2, "0")}`;
}

function formatAmount(value) {
  return value.toFixed(2);
}

function sumAmounts(transactions) {
  return transactions.reduce((sum, tx) => sum + tx.amount, 0);
}

function cleanText(value) {
  return String(value == null ? "" : value)
    .replace(/\s+/g, " ")
    .trim();
}

function cleanIban(value) {
  return cleanText(value).replace(/\s+/g, "").toUpperCase();
}

function requireText(value, label) {
  const text = cleanText(value);
  if (!text) {
    throw new Error(`"${label}" ontbreekt`);
  }
  return text;
}

function requireMaxLength(value, label, maxLength) {
  const text = String(value == null ? "" : value);
  if (!text) {
    throw new Error(`"${label}" ontbreekt`);
  }
  if (text.length > maxLength) {
    throw new Error(`"${label}" is langer dan ${maxLength} tekens`);
  }
  return text;
}

function enforceMaxLength(value, label, maxLength) {
  const text = String(value == null ? "" : value);
  if (text && text.length > maxLength) {
    throw new Error(`"${label}" is langer dan ${maxLength} tekens`);
  }
  return text;
}

function requireBicLength(value, label) {
  const bic = requireMaxLength(value, label, FIELD_LIMITS.bic);
  if (bic !== "NOTPROVIDED" && bic.length !== 8 && bic.length !== 11) {
    throw new Error(`"${label}" moet 8 of 11 tekens bevatten`);
  }
  return bic;
}

function formatCompactMessageId(date) {
  return `DD${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, "0")}${String(date.getDate()).padStart(2, "0")}${String(date.getHours()).padStart(2, "0")}${String(date.getMinutes()).padStart(2, "0")}${String(date.getSeconds()).padStart(2, "0")}`;
}

function xml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function escapeHtml(value) {
  return xml(value);
}

function setStatus(message, kind) {
  els.status.textContent = message;
  els.status.className = `message${kind === "error" ? " error" : kind === "warn" ? " warn" : ""}`;
  syncPreviewHeight();
}

function setSelectedFileName(name) {
  if (!els.selectedFileName) {
    return;
  }
  if (!name) {
    els.selectedFileName.textContent = "Geen bestand geselecteerd.";
    return;
  }
  els.selectedFileName.textContent = name;
}

async function handleSelectedFile(file) {
  resetDownload();
  setSelectedFileName(file.name);

  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, {
      defval: "",
      raw: false,
      dateNF: "dd-mm-yyyy",
    });

    hydrateFromRows(rows);
  } catch (error) {
    resetState();
    setStatus(
      error.message || "Het werkboek kon niet worden gelezen.",
      "error",
    );
  }
}

function syncPreviewHeight() {
  if (!els.inputCard || !els.previewCard || !els.previewTableWrap) {
    return;
  }

  const inputHeight = els.inputCard.offsetHeight;
  if (!inputHeight) {
    return;
  }

  els.previewCard.style.height = `${inputHeight}px`;

  const cardRect = els.previewCard.getBoundingClientRect();
  const tableRect = els.previewTableWrap.getBoundingClientRect();
  const cardStyles = getComputedStyle(els.previewCard);
  const bottomPadding = Number.parseFloat(cardStyles.paddingBottom) || 0;
  const availableHeight =
    inputHeight - (tableRect.top - cardRect.top) - bottomPadding;

  if (availableHeight > 0) {
    els.previewTableWrap.style.height = `${availableHeight}px`;
  }
}
