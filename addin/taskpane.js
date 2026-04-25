(function () {
  "use strict";

  const moduleFiles = [
    { prefix: "Controls", path: "../modules/controls.formula.txt" },
    { prefix: "get", path: "../modules/get.formula.txt" },
    { prefix: "kind", path: "../modules/kind.formula.txt" },
    { prefix: "CapitalPlanning", path: "../modules/capital_planning_report.formula.txt" },
    { prefix: "Analysis", path: "../modules/analysis.formula.txt" },
    { prefix: "defer", path: "../modules/defer.formula.txt" },
    { prefix: "Notes", path: "../modules/notes.formula.txt" },
    { prefix: "Phasing", path: "../modules/phasing.formula.txt" },
    { prefix: "Ready", path: "../modules/ready.formula.txt" },
    { prefix: "Search", path: "../modules/search.formula.txt" }
  ];

  const starterTables = [
    { sheet: "Planning Table", address: "A2", path: "../samples/planning_table_starter.tsv" },
    { sheet: "Cap Setup", address: "A2", path: "../samples/cap_setup_starter.tsv" }
  ];

  const reviewSheet = "Planning Review";
  const validationSheet = "Validation Lists";
  const requiredSheets = ["Planning Table", "Cap Setup", reviewSheet, validationSheet];
  const validationLists = {
    months: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
    groupFields: ["BU", "Region", "Site", "PM", "Category", "Revised Group", "Status", "Type"],
    futureFilters: ["All", "Exclude Future", "Keep F1 Only", "Keep F1+F2"],
    closedRows: ["SHOW", "HIDE"],
    statuses: ["Active", "Hold", "Closed", "In Service", "Skipping", "Canceled"],
    yesNo: ["Y", "N"]
  };
  const validationListColumns = [
    { key: "months", header: "Month" },
    { key: "groupFields", header: "Group Field" },
    { key: "futureFilters", header: "Future Filter" },
    { key: "closedRows", header: "Closed Rows" },
    { key: "statuses", header: "Status" },
    { key: "yesNo", header: "Yes No" }
  ];
  const visibleControlNames = [
    { name: "PM_Filter_Dropdowns", address: "B2", formula: "='Planning Review'!$B$2" },
    { name: "Future_Filter_Mode", address: "C2", formula: "='Planning Review'!$C$2" },
    { name: "HideClosed_Status", address: "D2", formula: "='Planning Review'!$D$2" },
    { name: "Burndown_Cut_Target", address: "E2", formula: "='Planning Review'!$E$2" }
  ];
  const demoOutputs = [
    {
      sheet: reviewSheet,
      title: "Capital Planning Report",
      formula: "=CapitalPlanning.CAPITAL_PLANNING_REPORT()",
      note: "Main report spill starts at A4."
    },
    {
      sheet: "BU Cap Scorecard",
      title: "BU Cap Scorecard",
      formula: "=Analysis.BU_CAP_SCORECARD()",
      note: "Cap and spend posture by BU."
    },
    {
      sheet: "Reforecast Queue",
      title: "Reforecast Queue",
      formula: "=Analysis.REFORECAST_QUEUE()",
      note: "Grouped action queue for forecast review."
    },
    {
      sheet: "PM Spend Report",
      title: "PM Spend Report",
      formula: "=Analysis.PM_SPEND_REPORT()",
      note: "Existing-work summary and job detail."
    },
    {
      sheet: "Working Budget",
      title: "Working Budget Screen",
      formula: "=Analysis.WORKING_BUDGET_SCREEN()",
      note: "Current-job screening before budget drafting."
    },
    {
      sheet: "Burndown",
      title: "Burndown Screen",
      formula: "=Analysis.BURNDOWN_SCREEN()",
      note: "Meeting view of remaining burn and drivers."
    }
  ];
  const requiredNames = [
    "PM_Filter_Dropdowns",
    "Future_Filter_Mode",
    "HideClosed_Status",
    "Burndown_Cut_Target",
    "Controls.PM_Filter_Dropdowns",
    "TRIMRANGE_KEEPBLANKS",
    "RBYROW",
    "get.TRIMRANGE_KEEPBLANKS",
    "get.GetFinanceBlock",
    "kind.RBYROW",
    "kind.CapByBU",
    "kind.PortfolioCap",
    "CapitalPlanning.CAPITAL_PLANNING_REPORT",
    "Analysis.PM_SPEND_REPORT",
    "Analysis.WORKING_BUDGET_SCREEN",
    "Analysis.BU_CAP_SCORECARD",
    "Analysis.REFORECAST_QUEUE",
    "Analysis.BURNDOWN_SCREEN"
  ];

  const logEl = document.getElementById("log");
  const buttons = Array.from(document.querySelectorAll("button"));

  Office.onReady((info) => {
    if (info.host !== Office.HostType.Excel) {
      writeLog("Open this add-in in Excel.");
      setButtons(false);
      return;
    }

    bind("setupWorkbook", setupWorkbook);
    bind("installModules", installModules);
    bind("validateWorkbook", validateWorkbook);
    bind("insertDemoOutputs", insertDemoOutputs);
    bind("runAll", runAll);
    writeLog("Ready.");
    setButtons(true);
  });

  function bind(id, fn) {
    document.getElementById(id).addEventListener("click", () => runGuarded(fn));
  }

  async function runGuarded(fn) {
    setButtons(false);
    try {
      await fn();
    } catch (error) {
      appendLog(`ERROR: ${error.message || error}`);
    } finally {
      setButtons(true);
    }
  }

  async function runAll() {
    clearLog();
    await setupWorkbook();
    await installModules();
    await validateWorkbook();
  }

  async function setupWorkbook() {
    appendLog("Creating starter sheets, controls, and validation lists...");
    const tables = [];
    for (const table of starterTables) {
      tables.push({
        sheet: table.sheet,
        address: table.address,
        values: parseTsv(await fetchText(table.path))
      });
    }

    await Excel.run(async (context) => {
      await ensureRequiredSheets(context);
      await context.sync();

      for (const table of tables) {
        const sheet = context.workbook.worksheets.getItem(table.sheet);
        const rowCount = table.values.length;
        const colCount = table.values[0].length;
        const range = sheet.getRange(table.address).getResizedRange(rowCount - 1, colCount - 1);
        range.values = table.values;
        range.format.autofitColumns();
      }

      buildValidationLists(context.workbook.worksheets.getItem(validationSheet));
      formatPlanningTable(context.workbook.worksheets.getItem("Planning Table"));
      formatCapSetup(context.workbook.worksheets.getItem("Cap Setup"));
      formatPlanningReview(context.workbook.worksheets.getItem(reviewSheet));
      await context.sync();
    });

    appendLog("Starter sheets, visible controls, dropdowns, and formats ready.");
  }

  async function installModules() {
    appendLog("Loading formula modules...");
    const formulas = [];
    const unqualifiedAliases = new Map();

    for (const moduleFile of moduleFiles) {
      const text = await fetchText(moduleFile.path);
      const parsed = parseFormulaModule(text);
      appendLog(`Parsed ${parsed.length} formulas from ${moduleFile.prefix}.`);

      for (const item of parsed) {
        formulas.push({
          name: `${moduleFile.prefix}.${item.name}`,
          formula: item.formula,
          comment: `Governed formula module: ${moduleFile.prefix}`
        });
        if (!unqualifiedAliases.has(item.name)) {
          unqualifiedAliases.set(item.name, moduleFile.prefix);
          formulas.push({
            name: item.name,
            formula: item.formula,
            comment: `Governed formula compatibility alias: ${moduleFile.prefix}`
          });
        } else {
          appendLog(
            `Skipped unqualified alias ${item.name}; already provided by ${unqualifiedAliases.get(item.name)}.`
          );
        }
      }
    }

    await Excel.run(async (context) => {
      for (const item of formulas) {
        await replaceName(context, item.name, item.formula, item.comment);
      }
      await bindVisibleControlNames(context);
      await context.sync();
    });

    appendLog(`Installed ${formulas.length} workbook names.`);
  }

  async function validateWorkbook() {
    appendLog("Validating workbook contract...");
    const expectedPlanningHeaders = parseTsv(await fetchText("../samples/planning_table_starter.tsv"))[0];
    const expectedCapHeaders = parseTsv(await fetchText("../samples/cap_setup_starter.tsv"))[0];

    const summary = await Excel.run(async (context) => {
      const sheets = {};
      for (const sheetName of requiredSheets) {
        const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
        sheet.load("name");
        sheets[sheetName] = sheet;
      }

      const names = {};
      for (const name of requiredNames) {
        const item = context.workbook.names.getItemOrNullObject(name);
        item.load("name");
        names[name] = item;
      }

      await context.sync();

      for (const sheetName of requiredSheets) {
        if (sheets[sheetName].isNullObject) {
          throw new Error(`Missing worksheet: ${sheetName}`);
        }
      }

      for (const name of requiredNames) {
        if (names[name].isNullObject) {
          throw new Error(`Missing workbook name: ${name}`);
        }
      }

      const controlNameItems = {};
      for (const control of visibleControlNames) {
        const item = context.workbook.names.getItemOrNullObject(control.name);
        item.load("name, formula");
        controlNameItems[control.name] = item;
      }

      const planning = context.workbook.worksheets.getItem("Planning Table");
      const capSetup = context.workbook.worksheets.getItem("Cap Setup");
      const review = context.workbook.worksheets.getItem(reviewSheet);

      const planningHeaders = planning.getRange("A2:BO2");
      const capHeaders = capSetup.getRange("A2:B2");
      const capRows = capSetup.getRange("A3:B100");
      const reviewControls = review.getRange("B2:E2");
      const reviewMonths = review.getRange("M2:N2");

      planningHeaders.load("values");
      capHeaders.load("values");
      capRows.load("values");
      reviewControls.load("values");
      reviewMonths.load("values");
      await context.sync();

      assertHeaderOrder(planningHeaders.values[0], expectedPlanningHeaders, "Planning Table");
      assertHeaderOrder(capHeaders.values[0], expectedCapHeaders, "Cap Setup");
      assertCapRowsAreValid(capRows.values);
      assertVisibleControls(reviewControls.values, reviewMonths.values);
      assertControlNamesBound(controlNameItems);

      return {
        sheetCount: requiredSheets.length,
        workbookNameCount: requiredNames.length,
        planningHeaderCount: expectedPlanningHeaders.length,
        capRowCount: countConfiguredCapRows(capRows.values),
        controlCount: visibleControlNames.length,
        dropdownListCount: validationListColumns.length
      };
    });

    appendLog("Workbook contract valid.");
    appendLog(renderValidationSummary(summary));
  }

  async function insertDemoOutputs() {
    appendLog("Validating before inserting demo outputs...");
    await validateWorkbook();
    appendLog("Inserting demo output formulas...");

    await Excel.run(async (context) => {
      await ensureSheets(context, unique(demoOutputs.map((output) => output.sheet)));
      await context.sync();

      for (const output of demoOutputs) {
        placeDemoOutput(context.workbook.worksheets.getItem(output.sheet), output);
      }

      context.workbook.worksheets.getItem(reviewSheet).activate();
      await context.sync();
    });

    appendLog(renderDemoOutputSummary());
  }

  async function ensureRequiredSheets(context) {
    await ensureSheets(context, requiredSheets);
  }

  async function ensureSheets(context, sheetNames) {
    const existingSheets = {};
    for (const sheetName of sheetNames) {
      const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
      sheet.load("name");
      existingSheets[sheetName] = sheet;
    }
    await context.sync();

    for (const sheetName of sheetNames) {
      if (existingSheets[sheetName].isNullObject) {
        context.workbook.worksheets.add(sheetName);
      }
    }
  }

  function placeDemoOutput(sheet, output) {
    if (output.sheet === reviewSheet) {
      sheet.getRange("A4").formulas = [[output.formula]];
      return;
    }

    sheet.getRange("A1:Z300").clear(Excel.ClearApplyTo.all);
    sheet.getRange("A1").values = [[output.title]];
    sheet.getRange("A2").values = [[output.note]];
    sheet.getRange("A4").formulas = [[output.formula]];
    sheet.getRange("A1").format.font.bold = true;
    sheet.getRange("A1").format.font.size = 16;
    sheet.getRange("A2").format.font.italic = true;
    sheet.getRange("A:Z").format.autofitColumns();
  }

  async function bindVisibleControlNames(context) {
    for (const control of visibleControlNames) {
      await replaceName(
        context,
        control.name,
        control.formula,
        `Governed formula visible control: ${reviewSheet}!${control.address}`
      );
    }
  }

  function buildValidationLists(sheet) {
    const rowCount = Math.max(...validationListColumns.map((column) => validationLists[column.key].length)) + 1;
    const values = Array.from({ length: rowCount }, (_, rowIndex) =>
      validationListColumns.map((column) => {
        if (rowIndex === 0) {
          return column.header;
        }
        return validationLists[column.key][rowIndex - 1] || "";
      })
    );

    const listRange = sheet.getRange("A1").getResizedRange(rowCount - 1, validationListColumns.length - 1);
    listRange.values = values;
    listRange.format.font.name = "Segoe UI";
    listRange.format.autofitColumns();
    sheet.getRange("A1:F1").format.font.bold = true;
    sheet.getRange("A1:F1").format.fill.color = "#D9EAF7";
  }

  function formatPlanningTable(sheet) {
    sheet.freezePanes.freezeRows(2);
    sheet.getRange("A2:BO2").format.font.bold = true;
    sheet.getRange("A2:BO2").format.wrapText = true;
    sheet.getRange("A2:BO2").format.fill.color = "#D9EAF7";
    for (const address of ["F2", "G2", "P2", "Q2", "BE2"]) {
      sheet.getRange(address).format.fill.color = "#FFF2CC";
    }
    applyNumberFormat(sheet.getRange("P3:BA234"), 232, 38, "$#,##0");
    applyNumberFormat(sheet.getRange("BM3:BM234"), 232, 1, "0");
    applyListValidation(sheet.getRange("F3:F234"), validationSource("E"));
    applyListValidation(sheet.getRange("M3:M234"), validationSource("F"));
    applyListValidation(sheet.getRange("O3:O234"), validationSource("F"));
    applyListValidation(sheet.getRange("BE3:BE234"), validationSource("F"));
    applyListValidation(sheet.getRange("BG3:BH234"), validationSource("F"));
    applyListValidation(sheet.getRange("BO3:BO234"), validationSource("F"));
    sheet.getRange("A:BO").format.autofitColumns();
  }

  function formatCapSetup(sheet) {
    sheet.freezePanes.freezeRows(2);
    sheet.getRange("A2:B2").format.font.bold = true;
    sheet.getRange("A2:B2").format.fill.color = "#D9EAF7";
    applyNumberFormat(sheet.getRange("B3:B100"), 98, 1, "$#,##0");
    applyNonNegativeValidation(sheet.getRange("B3:B100"));
    sheet.getRange("A:B").format.autofitColumns();
  }

  function formatPlanningReview(sheet) {
    sheet.freezePanes.freezeRows(3);
    sheet.getRange("A1:N3").clear(Excel.ClearApplyTo.all);
    sheet.getRange("J2:K6").clear(Excel.ClearApplyTo.all);
    sheet.getRange("A1").values = [["Planning Review"]];
    sheet.getRange("B1:E1").values = [["Group", "Future Filter", "Closed Rows", "Burndown Cut Target"]];
    sheet.getRange("A2:E2").values = [["Controls", "BU", "All", "SHOW", 0]];
    sheet.getRange("A3").values = [["Main report spill starts at A4. Columns O:R are reserved for notes."]];
    sheet.getRange("M1:N2").values = [
      ["Report As Of", "Defer As Of"],
      ["Mar", "Mar"]
    ];

    sheet.getRange("A1").format.font.bold = true;
    sheet.getRange("A1").format.font.size = 16;
    sheet.getRange("B1:E1").format.font.bold = true;
    sheet.getRange("M1:N1").format.font.bold = true;
    sheet.getRange("A2:E2").format.fill.color = "#F3F6FA";
    sheet.getRange("B2:E2").format.fill.color = "#FFF2CC";
    sheet.getRange("M2:N2").format.fill.color = "#FFF2CC";
    applyNumberFormat(sheet.getRange("E2"), 1, 1, "$#,##0");
    applyListValidation(sheet.getRange("B2"), validationSource("B"));
    applyListValidation(sheet.getRange("C2"), validationSource("C"));
    applyListValidation(sheet.getRange("D2"), validationSource("D"));
    applyListValidation(sheet.getRange("M2:N2"), validationSource("A"));
    applyNonNegativeValidation(sheet.getRange("E2"));
    sheet.getRange("A:N").format.autofitColumns();
  }

  function validationSource(columnLetter) {
    const index = columnLetter.charCodeAt(0) - "A".charCodeAt(0);
    const listKey = validationListColumns[index].key;
    const endRow = validationLists[listKey].length + 1;
    return `='${validationSheet}'!$${columnLetter}$2:$${columnLetter}$${endRow}`;
  }

  function applyListValidation(range, source) {
    range.dataValidation.clear();
    range.dataValidation.rule = {
      list: {
        inCellDropDown: true,
        source
      }
    };
  }

  function applyNonNegativeValidation(range) {
    range.dataValidation.clear();
    range.dataValidation.rule = {
      decimal: {
        formula1: "0",
        operator: Excel.DataValidationOperator.greaterThanOrEqualTo
      }
    };
  }

  function applyNumberFormat(range, rowCount, columnCount, format) {
    range.numberFormat = Array.from({ length: rowCount }, () => Array(columnCount).fill(format));
  }

  function assertHeaderOrder(actualHeaders, expectedHeaders, label) {
    const actual = actualHeaders.map((value) => String(value).trim());
    if (actual.length !== expectedHeaders.length) {
      throw new Error(`${label} header count is ${actual.length}; expected ${expectedHeaders.length}.`);
    }
    expectedHeaders.forEach((expected, index) => {
      if (actual[index] !== expected) {
        throw new Error(`${label} header ${index + 1} should be ${expected}; found ${actual[index] || "(blank)"}.`);
      }
    });
  }

  function assertCapRowsAreValid(rows) {
    rows.forEach((row, index) => {
      const bu = String(row[0] || "").trim();
      const cap = row[1];
      if (bu && !(Number(cap) >= 0)) {
        throw new Error(`Cap Setup row ${index + 3} has BU but Cap is not a non-negative number.`);
      }
    });
  }

  function countConfiguredCapRows(rows) {
    return rows.filter((row) => String(row[0] || "").trim()).length;
  }

  function assertVisibleControls(controlValues, monthValues) {
    assertAllowed("PM_Filter_Dropdowns", controlValues[0][0], validationLists.groupFields);
    assertAllowed("Future_Filter_Mode", controlValues[0][1], validationLists.futureFilters);
    assertAllowed("HideClosed_Status", controlValues[0][2], validationLists.closedRows);
    if (!(Number(controlValues[0][3]) >= 0)) {
      throw new Error("Burndown_Cut_Target control must be a non-negative number.");
    }
    assertAllowed("Planning Review!M2", monthValues[0][0], validationLists.months);
    assertAllowed("Planning Review!N2", monthValues[0][1], validationLists.months);
  }

  function assertAllowed(label, value, allowedValues) {
    const normalized = String(value || "").trim().toLowerCase();
    const allowed = new Set(allowedValues.map((item) => item.toLowerCase()));
    if (!allowed.has(normalized)) {
      throw new Error(`${label} value ${value || "(blank)"} is not allowed.`);
    }
  }

  function assertControlNamesBound(controlNameItems) {
    for (const control of visibleControlNames) {
      const item = controlNameItems[control.name];
      if (item.isNullObject) {
        throw new Error(`Missing workbook name: ${control.name}`);
      }
      if (normalizeFormula(item.formula) !== normalizeFormula(control.formula)) {
        throw new Error(`${control.name} should point to ${control.formula}; found ${item.formula}.`);
      }
    }
  }

  function normalizeFormula(formula) {
    return String(formula || "").replace(/\s/g, "").toUpperCase();
  }

  function renderValidationSummary(summary) {
    return [
      "Validation summary:",
      `- Sheets present: ${summary.sheetCount}/${requiredSheets.length}`,
      `- Workbook names installed: ${summary.workbookNameCount}/${requiredNames.length}`,
      `- Planning Table headers: ${summary.planningHeaderCount}`,
      `- Cap Setup rows with BU: ${summary.capRowCount}`,
      `- Visible controls bound: ${summary.controlCount}/${visibleControlNames.length}`,
      `- Dropdown lists ready: ${summary.dropdownListCount}`
    ].join("\n");
  }

  function renderDemoOutputSummary() {
    return [
      "Demo outputs inserted:",
      ...demoOutputs.map((output) => `- ${output.sheet}: A4 -> ${output.formula}`)
    ].join("\n");
  }

  function unique(values) {
    return Array.from(new Set(values));
  }

  async function replaceName(context, name, formula, comment) {
    const existing = context.workbook.names.getItemOrNullObject(name);
    existing.load("name");
    await context.sync();
    if (!existing.isNullObject) {
      existing.delete();
      await context.sync();
    }
    context.workbook.names.add(name, formula, comment);
  }

  async function fetchText(path) {
    const response = await fetch(path, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`Unable to load ${path}: ${response.status}`);
    }
    return response.text();
  }

  function parseTsv(text) {
    const rows = text.replace(/\r/g, "").split("\n").filter(Boolean);
    return rows.map((row) => row.split("\t"));
  }

  function parseFormulaModule(text) {
    const source = text.replace(/\r/g, "");
    const matches = Array.from(source.matchAll(/^([A-Za-z_][A-Za-z0-9_]*)\s*=/gm));
    return matches.map((match, index) => {
      const name = match[1];
      const bodyStart = match.index + match[0].length;
      const bodyEnd = index + 1 < matches.length ? matches[index + 1].index : source.length;
      let body = source.slice(bodyStart, bodyEnd).trim();
      if (body.endsWith(";")) {
        body = body.slice(0, -1).trim();
      }
      return {
        name,
        formula: `=${stripBlockComments(body)}`
      };
    });
  }

  function stripBlockComments(text) {
    return text.replace(/\/\*[\s\S]*?\*\//g, "");
  }

  function setButtons(enabled) {
    buttons.forEach((button) => {
      button.disabled = !enabled;
    });
  }

  function clearLog() {
    logEl.textContent = "";
  }

  function writeLog(message) {
    logEl.textContent = message;
  }

  function appendLog(message) {
    logEl.textContent = `${logEl.textContent ? `${logEl.textContent}\n` : ""}${message}`;
  }
})();
