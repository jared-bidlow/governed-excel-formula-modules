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
  const requiredSheets = ["Planning Table", "Cap Setup", reviewSheet];
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
    "Analysis.BU_CAP_SCORECARD",
    "Analysis.REFORECAST_QUEUE"
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
    appendLog("Creating starter sheets...");
    const tables = [];
    for (const table of starterTables) {
      tables.push({
        sheet: table.sheet,
        address: table.address,
        values: parseTsv(await fetchText(table.path))
      });
    }

    await Excel.run(async (context) => {
      const existingSheets = {};
      for (const sheetName of requiredSheets) {
        const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
        sheet.load("name");
        existingSheets[sheetName] = sheet;
      }
      await context.sync();

      for (const sheetName of requiredSheets) {
        if (existingSheets[sheetName].isNullObject) {
          context.workbook.worksheets.add(sheetName);
        }
      }
      await context.sync();

      for (const table of tables) {
        const sheet = context.workbook.worksheets.getItem(table.sheet);
        const rowCount = table.values.length;
        const colCount = table.values[0].length;
        const range = sheet.getRange(table.address).getResizedRange(rowCount - 1, colCount - 1);
        range.values = table.values;
        range.format.autofitColumns();
      }

      context.workbook.worksheets.getItem(reviewSheet).getRange("M2").values = [["Mar"]];
      await context.sync();
    });

    appendLog("Starter sheets ready.");
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
      await context.sync();
    });

    appendLog(`Installed ${formulas.length} workbook names.`);
  }

  async function validateWorkbook() {
    appendLog("Validating workbook contract...");
    await Excel.run(async (context) => {
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

      const headers = context.workbook.worksheets
        .getItem("Planning Table")
        .getRange("A2:BO2");
      headers.load("values");
      await context.sync();

      const headerSet = new Set(headers.values[0].map((value) => String(value).trim()));
      for (const header of ["BU", "Annual Projected", "Current Authorized Amount"]) {
        if (!headerSet.has(header)) {
          throw new Error(`Missing Planning Table header: ${header}`);
        }
      }
    });

    appendLog("Workbook contract valid.");
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
