(function () {
  "use strict";

  const applicationData = {
    sheets: {
      planningTable: "Planning Table",
      capSetup: "Cap Setup",
      planningReview: "Planning Review",
      validationLists: "Validation Lists",
      notesStaging: "Decision Staging",
      assetSetup: "Asset Setup",
      assetRegister: "Asset Register",
      projectAssetMap: "Project Asset Map",
      semanticAssets: "Semantic Assets",
      assetChanges: "Asset Changes",
      assetStateHistory: "Asset State History"
    },
    planningTable: {
      headerRange: "A2:BL2",
      headerRow: 2,
      dataStartRow: 3,
      maxValidationRow: 2000,
      requiredHeaderFill: ["F2", "G2", "O2", "P2", "BE2"],
      numberFormats: [
        { address: "O3:AZ234", rows: 232, columns: 38, format: "$#,##0" },
        { address: "BJ3:BJ234", rows: 232, columns: 1, format: "0" }
      ]
    },
    capSetup: {
      headerRange: "A2:B2",
      dataRange: "A3:B100",
      capRange: "B3:B100"
    },
    starterTables: [
      { sheet: "Planning Table", address: "A2", path: "../samples/planning_table_starter.tsv" },
      { sheet: "Cap Setup", address: "A2", path: "../samples/cap_setup_starter.tsv" }
    ],
    notesWorkflow: {
      tableName: "tblDecisionStaging",
      tableAddress: "A1",
      noteHeaders: ["ExistingMeetingNotes", "NewPlanningNotes", "NewTimeline", "NewStatus"],
      stagingHeaders: [
        "GroupType",
        "GroupValue",
        "Category",
        "ProjDesc",
        "AnnualProj",
        "ActualsYTD",
        "ExistingMeetingNotes",
        "NewPlanningNotes",
        "NewTimeline",
        "NewStatus",
        "ApplyAction",
        "PlanningNotes_New",
        "Timeline_New",
        "Comments_New",
        "Status_New",
        "KeyStatus",
        "BudgetMatchCount",
        "ApplyReady",
        "ApplyStatus",
        "AppliedOn",
        "ApplyMessage",
        "BudgetRowFound"
      ]
    },
    assetWorkflow: {
      tables: [
        {
          sheet: "Asset Register",
          tableName: "tblAssets",
          address: "A1",
          headers: [
            "AssetID",
            "AssetName",
            "AssetType",
            "Site",
            "Location",
            "Department",
            "Owner",
            "Status",
            "InServiceDate",
            "Condition",
            "Criticality",
            "ReplacementCost",
            "UsefulLifeYears",
            "LastInspectionDate",
            "NextReviewDate",
            "LinkedProjectID",
            "Notes"
          ]
        },
        {
          sheet: "Semantic Assets",
          tableName: "tblSemanticAssets",
          address: "A1",
          headers: [
            "ProjectKey",
            "ProjectDescription",
            "CandidateAssetId",
            "AssetLabel",
            "AssetType",
            "ProposedChangeType",
            "SourceAssetId",
            "TargetAssetId",
            "InstalledState",
            "EvidenceId",
            "PromotionStatus",
            "ApplyReady",
            "ApplyStatus",
            "AppliedOn",
            "ApplyMessage"
          ]
        },
        {
          sheet: "Asset Setup",
          tableName: "tblAssetPromotionQueue",
          address: "A1",
          headers: [
            "ProjectKey",
            "CandidateAssetId",
            "ProjectDescription",
            "AssetLabel",
            "AssetType",
            "ProposedChangeType",
            "SourceAssetId",
            "TargetAssetId",
            "InstalledState",
            "EvidenceId",
            "PromotionStatus",
            "ApplyReady",
            "ApplyStatus",
            "AppliedOn",
            "ApplyMessage"
          ]
        },
        {
          sheet: "Asset Setup",
          tableName: "tblAssetMappingStaging",
          address: "A6",
          headers: [
            "ProjectKey",
            "ProjectDescription",
            "ChangeType",
            "SourceAssetId",
            "TargetAssetId",
            "AssetId",
            "AssetLabel",
            "AssetType",
            "InstalledState",
            "EvidenceId",
            "MappingStatus",
            "ApplyReady",
            "ApplyStatus",
            "AppliedOn",
            "ApplyMessage"
          ]
        },
        {
          sheet: "Project Asset Map",
          tableName: "tblProjectAssetMap",
          address: "A1",
          headers: [
            "ProjectKey",
            "ProjectDescription",
            "AssetId",
            "AssetLabel",
            "AssetType",
            "AssetState",
            "EvidenceId",
            "MappingStatus",
            "ApplyStatus",
            "AppliedOn",
            "ApplyMessage"
          ]
        },
        {
          sheet: "Asset Changes",
          tableName: "tblAssetChanges",
          address: "A1",
          headers: [
            "ChangeId",
            "ProjectKey",
            "ChangeType",
            "SourceAssetId",
            "TargetAssetId",
            "InstalledState",
            "EvidenceId",
            "ChangeStatus",
            "AppliedOn",
            "ApplyMessage"
          ]
        },
        {
          sheet: "Asset State History",
          tableName: "tblAssetStateHistory",
          address: "A1",
          headers: [
            "EventId",
            "AssetId",
            "ProjectKey",
            "AssetState",
            "EvidenceId",
            "EventSource",
            "EventOn",
            "ApplyMessage"
          ]
        }
      ],
      relationshipLists: [
        {
          key: "assetIds",
          header: "Asset ID",
          formula:
            '=LET(ids,VSTACK(tblAssets[AssetID],tblSemanticAssets[CandidateAssetId],tblSemanticAssets[SourceAssetId],tblSemanticAssets[TargetAssetId],tblAssetPromotionQueue[CandidateAssetId],tblAssetPromotionQueue[SourceAssetId],tblAssetPromotionQueue[TargetAssetId],tblAssetMappingStaging[SourceAssetId],tblAssetMappingStaging[TargetAssetId],tblAssetMappingStaging[AssetId],tblProjectAssetMap[AssetId],tblAssetChanges[SourceAssetId],tblAssetChanges[TargetAssetId],tblAssetStateHistory[AssetId]),IFERROR(SORT(UNIQUE(FILTER(ids,ids<>""))),""))'
        },
        {
          key: "projectKeys",
          header: "Project Key",
          formula:
            '=LET(keys,VSTACK(tblAssets[LinkedProjectID],tblSemanticAssets[ProjectKey],tblAssetPromotionQueue[ProjectKey],tblAssetMappingStaging[ProjectKey],tblProjectAssetMap[ProjectKey],tblAssetChanges[ProjectKey],tblAssetStateHistory[ProjectKey]),IFERROR(SORT(UNIQUE(FILTER(keys,keys<>""))),""))'
        }
      ],
      tableValidationRules: {
        tblAssets: [
          { header: "Status", listKey: "assetStatuses" },
          { header: "Condition", listKey: "assetConditions" },
          { header: "Criticality", listKey: "assetCriticalities" },
          { header: "LinkedProjectID", relationshipListKey: "projectKeys" }
        ],
        tblSemanticAssets: [
          { header: "ProjectKey", relationshipListKey: "projectKeys" },
          { header: "CandidateAssetId", relationshipListKey: "assetIds" },
          { header: "ProposedChangeType", listKey: "assetChangeTypes" },
          { header: "SourceAssetId", relationshipListKey: "assetIds" },
          { header: "TargetAssetId", relationshipListKey: "assetIds" },
          { header: "InstalledState", listKey: "assetStates" },
          { header: "PromotionStatus", listKey: "assetPromotionStatuses" },
          { header: "ApplyReady", listKey: "yesNo" }
        ],
        tblAssetPromotionQueue: [
          { header: "ProjectKey", relationshipListKey: "projectKeys" },
          { header: "CandidateAssetId", relationshipListKey: "assetIds" },
          { header: "ProposedChangeType", listKey: "assetChangeTypes" },
          { header: "SourceAssetId", relationshipListKey: "assetIds" },
          { header: "TargetAssetId", relationshipListKey: "assetIds" },
          { header: "InstalledState", listKey: "assetStates" },
          { header: "PromotionStatus", listKey: "assetPromotionStatuses" },
          { header: "ApplyReady", listKey: "yesNo" }
        ],
        tblAssetMappingStaging: [
          { header: "ProjectKey", relationshipListKey: "projectKeys" },
          { header: "ChangeType", listKey: "assetChangeTypes" },
          { header: "SourceAssetId", relationshipListKey: "assetIds" },
          { header: "TargetAssetId", relationshipListKey: "assetIds" },
          { header: "AssetId", relationshipListKey: "assetIds" },
          { header: "InstalledState", listKey: "assetStates" },
          { header: "MappingStatus", listKey: "assetMappingStatuses" },
          { header: "ApplyReady", listKey: "yesNo" }
        ],
        tblProjectAssetMap: [
          { header: "ProjectKey", relationshipListKey: "projectKeys" },
          { header: "AssetId", relationshipListKey: "assetIds" },
          { header: "AssetState", listKey: "assetStates" },
          { header: "MappingStatus", listKey: "assetMappingStatuses" }
        ],
        tblAssetChanges: [
          { header: "ProjectKey", relationshipListKey: "projectKeys" },
          { header: "ChangeType", listKey: "assetChangeTypes" },
          { header: "SourceAssetId", relationshipListKey: "assetIds" },
          { header: "TargetAssetId", relationshipListKey: "assetIds" },
          { header: "InstalledState", listKey: "assetStates" },
          { header: "ChangeStatus", listKey: "assetChangeStatuses" }
        ],
        tblAssetStateHistory: [
          { header: "AssetId", relationshipListKey: "assetIds" },
          { header: "ProjectKey", relationshipListKey: "projectKeys" },
          { header: "AssetState", listKey: "assetStates" }
        ]
      }
    },
    moduleFiles: [
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
    ],
    dropdownLists: {
      months: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
      groupFields: ["Revised Group", "Site", "Region", "PM", "BU", "Category"],
      futureFilters: ["All", "Exclude Future", "Keep F1 Only", "Keep F1+F2"],
      closedRows: ["SHOW", "HIDE"],
      statuses: ["Active", "Hold", "Closed", "In Service", "Skipping", "Canceled"],
      yesNo: ["Y", "N"],
      assetStatuses: ["planned", "active", "in_service", "maintenance", "retired"],
      assetConditions: ["new", "good", "fair", "poor", "critical"],
      assetCriticalities: ["low", "medium", "high", "critical"],
      assetChangeTypes: ["new_asset", "replace_asset", "upgrade_asset"],
      assetStates: ["mapped", "planned", "installed", "retired"],
      assetPromotionStatuses: ["draft", "review", "accepted", "ready", "project_ready", "rejected"],
      assetMappingStatuses: ["draft", "active", "ready", "needs_review", "inactive"],
      assetChangeStatuses: ["draft", "ready", "applied", "needs_review", "blocked"]
    },
    validationListColumns: [
      { key: "months", header: "Month" },
      { key: "groupFields", header: "Group Field" },
      { key: "futureFilters", header: "Future Filter" },
      { key: "closedRows", header: "Closed Rows" },
      { key: "statuses", header: "Status" },
      { key: "yesNo", header: "Yes No" },
      { key: "assetStatuses", header: "Asset Status" },
      { key: "assetConditions", header: "Asset Condition" },
      { key: "assetCriticalities", header: "Asset Criticality" },
      { key: "assetChangeTypes", header: "Asset Change Type" },
      { key: "assetStates", header: "Asset State" },
      { key: "assetPromotionStatuses", header: "Asset Promotion Status" },
      { key: "assetMappingStatuses", header: "Asset Mapping Status" },
      { key: "assetChangeStatuses", header: "Asset Change Status" }
    ],
    visibleControls: [
      { name: "PM_Filter_Dropdowns", address: "B2", formula: "='Planning Review'!$B$2" },
      { name: "Future_Filter_Mode", address: "C2", formula: "='Planning Review'!$C$2" },
      { name: "HideClosed_Status", address: "D2", formula: "='Planning Review'!$D$2" },
      { name: "Burndown_Cut_Target", address: "E2", formula: "='Planning Review'!$E$2" }
    ],
    rowValidationRules: [
      { sheet: "Planning Table", header: "Status", listKey: "statuses" },
      {
        sheet: "Planning Table",
        header: "Chargeable",
        listKey: "yesNo",
        purpose: "Internal labor chargeability flag used by Search and Ready export helpers."
      },
      { sheet: "Planning Table", header: "Internal Eligible", listKey: "yesNo" },
      { sheet: "Planning Table", header: "Canceled", listKey: "yesNo" }
    ],
    demoOutputs: [
      {
        sheet: "Planning Review",
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
      },
      {
        sheet: "Internal Jobs",
        title: "Internal Jobs Export",
        formula: "=Ready.InternalJobs_Export()",
        note: "Header-driven internal work export for readiness smoke testing."
      }
    ],
    requiredNames: [
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
      "Analysis.BURNDOWN_SCREEN",
      "Ready.ColumnOrBlank",
      "Ready.InternalEligible",
      "Ready.ChargeableFlag",
      "Ready.InternalReady3",
      "Ready.InternalJobs_Export"
    ]
  };

  const moduleFiles = applicationData.moduleFiles;
  const starterTables = applicationData.starterTables;
  const reviewSheet = applicationData.sheets.planningReview;
  const validationSheet = applicationData.sheets.validationLists;
  const requiredSheets = [
    applicationData.sheets.planningTable,
    applicationData.sheets.capSetup,
    applicationData.sheets.planningReview,
    applicationData.sheets.validationLists
  ];
  const notesWorkflow = applicationData.notesWorkflow;
  const assetWorkflow = applicationData.assetWorkflow;
  const validationLists = applicationData.dropdownLists;
  const validationListColumns = applicationData.validationListColumns;
  const visibleControlNames = applicationData.visibleControls;
  const rowValidationRules = applicationData.rowValidationRules;
  const demoOutputs = applicationData.demoOutputs;
  const requiredNames = applicationData.requiredNames;

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
    bind("setupNotesWorkflow", setupNotesWorkflow);
    bind("setupAssetWorkflow", setupAssetWorkflow);
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
    await insertDemoOutputs({ validateFirst: false });
    await setupNotesWorkflow();
    appendLog("Standard setup complete. Asset workflow remains optional; run Setup Asset Workflow only when asset tables are wanted.");
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
      formatPlanningTable(
        context.workbook.worksheets.getItem(applicationData.sheets.planningTable),
        starterHeadersFor(tables, applicationData.sheets.planningTable)
      );
      formatCapSetup(context.workbook.worksheets.getItem(applicationData.sheets.capSetup));
      formatPlanningReview(context.workbook.worksheets.getItem(reviewSheet));
      await context.sync();
    });

    appendLog("Starter sheets, visible controls, dropdowns, and formats ready.");
  }

  async function setupNotesWorkflow() {
    appendLog("Setting up notes workflow...");

    await Excel.run(async (context) => {
      await ensureSheets(context, [reviewSheet, validationSheet, applicationData.sheets.notesStaging]);
      await context.sync();

      const review = context.workbook.worksheets.getItem(reviewSheet);
      buildValidationLists(context.workbook.worksheets.getItem(validationSheet));
      const staging = context.workbook.worksheets.getItem(applicationData.sheets.notesStaging);
      formatPlanningReviewNotes(review);
      await refreshTableFromHeaders(
        context,
        staging,
        notesWorkflow.tableName,
        notesWorkflow.tableAddress,
        notesWorkflow.stagingHeaders
      );
      formatWorkflowSheet(staging, notesWorkflow.stagingHeaders.length);
      await context.sync();
    });

    appendLog("Notes workflow ready: Planning Review notes columns and Decision Staging table are configured.");
  }

  async function setupAssetWorkflow() {
    appendLog("Setting up optional asset workflow...");

    await Excel.run(async (context) => {
      await ensureSheets(context, unique([...assetWorkflow.tables.map((table) => table.sheet), validationSheet]));
      await context.sync();

      buildValidationLists(context.workbook.worksheets.getItem(validationSheet));
      const createdTables = [];
      for (const table of assetWorkflow.tables) {
        const sheet = context.workbook.worksheets.getItem(table.sheet);
        const workbookTable = await refreshTableFromHeaders(context, sheet, table.tableName, table.address, table.headers);
        formatWorkflowSheet(sheet, table.headers.length);
        createdTables.push({ definition: table, workbookTable });
      }

      buildAssetRelationshipLists(context.workbook.worksheets.getItem(validationSheet));
      for (const item of createdTables) {
        applyTableValidationRules(
          item.workbookTable,
          assetWorkflow.tableValidationRules[item.definition.tableName] || []
        );
      }

      await context.sync();
    });

    appendLog("Optional asset workflow ready. Asset setup is opt-in and was not added to the normal run.");
  }

  function starterHeadersFor(tables, sheetName) {
    const table = tables.find((item) => item.sheet === sheetName);
    if (!table || !table.values.length) {
      throw new Error(`Missing starter table headers for ${sheetName}.`);
    }
    return table.values[0];
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
    const expectedPlanningHeaders = parseTsv(
      await fetchText(starterTables.find((table) => table.sheet === applicationData.sheets.planningTable).path)
    )[0];
    const expectedCapHeaders = parseTsv(
      await fetchText(starterTables.find((table) => table.sheet === applicationData.sheets.capSetup).path)
    )[0];

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

      const planning = context.workbook.worksheets.getItem(applicationData.sheets.planningTable);
      const capSetup = context.workbook.worksheets.getItem(applicationData.sheets.capSetup);
      const review = context.workbook.worksheets.getItem(reviewSheet);

      const planningHeaders = planning.getRange(applicationData.planningTable.headerRange);
      const capHeaders = capSetup.getRange(applicationData.capSetup.headerRange);
      const capRows = capSetup.getRange(applicationData.capSetup.dataRange);
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
      assertRowValidationRulesConfigured(planningHeaders.values[0]);
      assertCapRowsAreValid(capRows.values);
      assertVisibleControls(reviewControls.values, reviewMonths.values);
      assertControlNamesBound(controlNameItems);

      return {
        sheetCount: requiredSheets.length,
        workbookNameCount: requiredNames.length,
        planningHeaderCount: expectedPlanningHeaders.length,
        capRowCount: countConfiguredCapRows(capRows.values),
        controlCount: visibleControlNames.length,
        dropdownListCount: validationListColumns.length,
        rowValidationRuleCount: rowValidationRules.length
      };
    });

    appendLog("Workbook contract valid.");
    appendLog(renderValidationSummary(summary));
  }

  async function insertDemoOutputs(options = {}) {
    const validateFirst = options.validateFirst !== false;
    if (validateFirst) {
      appendLog("Validating before inserting demo outputs...");
      await validateWorkbook();
    }
    appendLog("Inserting demo output formulas...");

    await Excel.run(async (context) => {
      await ensureSheets(context, unique(demoOutputs.map((output) => output.sheet)));
      await context.sync();

      const mainOutput = demoOutputs.find((output) => output.sheet === reviewSheet);
      const review = context.workbook.worksheets.getItem(reviewSheet);
      const mainSpillRange = review.getRange("A4:N200");
      mainSpillRange.load(["values", "formulas"]);
      await context.sync();
      assertMainReportSpillReady(mainSpillRange.values, mainSpillRange.formulas, mainOutput.formula);

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
    const headerRange = sheet.getRange("A1").getResizedRange(0, validationListColumns.length - 1);
    headerRange.format.font.bold = true;
    headerRange.format.fill.color = "#D9EAF7";
  }

  function buildAssetRelationshipLists(sheet) {
    const startColumnIndex = validationListColumns.length + 1;
    for (let index = 0; index < assetWorkflow.relationshipLists.length; index += 1) {
      const list = assetWorkflow.relationshipLists[index];
      const column = columnName(startColumnIndex + index);
      sheet.getRange(`${column}1`).values = [[list.header]];
      sheet.getRange(`${column}2`).formulas = [[list.formula]];
      sheet.getRange(`${column}1`).format.font.bold = true;
      sheet.getRange(`${column}1`).format.fill.color = "#D9EAF7";
    }
    const lastColumn = columnName(startColumnIndex + assetWorkflow.relationshipLists.length - 1);
    sheet.getRange(`${columnName(startColumnIndex)}:${lastColumn}`).format.autofitColumns();
  }

  async function refreshTableFromHeaders(context, sheet, tableName, address, headers) {
    const existing = context.workbook.tables.getItemOrNullObject(tableName);
    existing.load("name");
    await context.sync();

    if (!existing.isNullObject) {
      existing.delete();
      await context.sync();
    }

    const values = [headers, Array(headers.length).fill("")];
    const range = sheet.getRange(address).getResizedRange(1, headers.length - 1);
    range.values = values;
    const table = sheet.tables.add(range, true);
    table.name = tableName;
    table.style = "TableStyleMedium2";
    table.showFilterButton = true;
    const headerRange = table.getHeaderRowRange();
    headerRange.format.font.bold = true;
    headerRange.format.font.color = "#000000";
    headerRange.format.fill.color = "#D9EAF7";
    return table;
  }

  function formatPlanningReviewNotes(sheet) {
    sheet.getRange("O4:R4").values = [notesWorkflow.noteHeaders];
    sheet.getRange("O4:R4").format.font.bold = true;
    sheet.getRange("O4:R4").format.fill.color = "#D9EAF7";
    sheet.getRange("O5:O200").clear(Excel.ClearApplyTo.all);
    sheet.getRange("O5").formulas = [["=IFERROR(Notes.Existing,\"\")"]];
    sheet.getRange("O5:O200").format.fill.color = "#F3F6FA";
    sheet.getRange("P5:R200").format.fill.color = "#FFF2CC";
    sheet.getRange("O:O").format.wrapText = false;
    sheet.getRange("P:R").format.wrapText = true;
    try {
      applyListValidation(sheet.getRange("R5:R200"), validationSourceForList("statuses"));
    } catch (error) {
      appendLog("Skipped NewStatus dropdown because validation lists are not ready.");
    }
    sheet.getRange("O:R").format.autofitColumns();
  }

  function formatWorkflowSheet(sheet, headerCount) {
    const lastColumn = columnName(headerCount);
    sheet.freezePanes.freezeRows(1);
    const headerRange = sheet.getRange(`A1:${lastColumn}1`);
    headerRange.format.font.bold = true;
    headerRange.format.font.color = "#000000";
    headerRange.format.fill.color = "#D9EAF7";
    sheet.getRange(`A:${lastColumn}`).format.wrapText = true;
    sheet.getRange(`A:${lastColumn}`).format.autofitColumns();
  }

  function formatPlanningTable(sheet, headers) {
    const table = applicationData.planningTable;
    sheet.freezePanes.freezeRows(table.headerRow);
    sheet.getRange(table.headerRange).format.font.bold = true;
    sheet.getRange(table.headerRange).format.wrapText = true;
    sheet.getRange(table.headerRange).format.fill.color = "#D9EAF7";
    for (const address of table.requiredHeaderFill) {
      sheet.getRange(address).format.fill.color = "#FFF2CC";
    }
    for (const numberFormat of table.numberFormats) {
      applyNumberFormat(
        sheet.getRange(numberFormat.address),
        numberFormat.rows,
        numberFormat.columns,
        numberFormat.format
      );
    }
    applyRowValidationRules(sheet, headers);
    sheet.getRange("A:BL").format.autofitColumns();
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
    applyListValidation(sheet.getRange("B2"), validationSourceForList("groupFields"));
    applyListValidation(sheet.getRange("C2"), validationSourceForList("futureFilters"));
    applyListValidation(sheet.getRange("D2"), validationSourceForList("closedRows"));
    applyListValidation(sheet.getRange("M2:N2"), validationSourceForList("months"));
    applyNonNegativeValidation(sheet.getRange("E2"));
    sheet.getRange("A:N").format.autofitColumns();
  }

  function applyRowValidationRules(sheet, headers) {
    for (const rule of rowValidationRulesFor(applicationData.sheets.planningTable)) {
      const address = dataRangeForHeader(
        headers,
        rule.header,
        applicationData.planningTable.dataStartRow,
        applicationData.planningTable.maxValidationRow
      );
      applyListValidation(sheet.getRange(address), validationSourceForList(rule.listKey));
    }
  }

  function applyTableValidationRules(table, rules) {
    for (const rule of rules) {
      const source = validationSourceForRule(rule);
      const range = table.columns.getItem(rule.header).getDataBodyRange();
      applyListValidation(range, source, { allowUnknown: Boolean(rule.relationshipListKey) });
    }
  }

  function dataRangeForHeader(headers, header, startRow, endRow) {
    const index = headerIndex(headers, header);
    const column = columnName(index + 1);
    return `${column}${startRow}:${column}${endRow}`;
  }

  function rowValidationRulesFor(sheetName) {
    return rowValidationRules.filter((rule) => rule.sheet === sheetName);
  }

  function validationSourceForList(listKey) {
    const index = validationListColumns.findIndex((column) => column.key === listKey);
    if (index < 0) {
      throw new Error(`Unknown validation list: ${listKey}`);
    }
    const columnLetter = columnName(index + 1);
    const endRow = validationLists[listKey].length + 1;
    return `='${validationSheet}'!$${columnLetter}$2:$${columnLetter}$${endRow}`;
  }

  function validationSourceForRule(rule) {
    if (rule.listKey) {
      return validationSourceForList(rule.listKey);
    }
    if (rule.relationshipListKey) {
      return validationSourceForRelationshipList(rule.relationshipListKey);
    }
    throw new Error(`Validation rule for ${rule.header} has no source list.`);
  }

  function validationSourceForRelationshipList(listKey) {
    const index = assetWorkflow.relationshipLists.findIndex((list) => list.key === listKey);
    if (index < 0) {
      throw new Error(`Unknown asset relationship list: ${listKey}`);
    }
    const columnLetter = columnName(validationListColumns.length + index + 1);
    return `='${validationSheet}'!$${columnLetter}$2#`;
  }

  function applyListValidation(range, source, options = {}) {
    range.dataValidation.clear();
    range.dataValidation.rule = {
      list: {
        inCellDropDown: true,
        source
      }
    };
    if (options.allowUnknown) {
      range.dataValidation.errorAlert = {
        showAlert: false
      };
    }
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

  function assertRowValidationRulesConfigured(headers) {
    for (const rule of rowValidationRulesFor(applicationData.sheets.planningTable)) {
      headerIndex(headers, rule.header);
      if (!validationLists[rule.listKey]) {
        throw new Error(`Row validation for ${rule.header} uses unknown list ${rule.listKey}.`);
      }
    }
  }

  function headerIndex(headers, header) {
    const actual = headers.map((value) => String(value || "").trim());
    const index = actual.indexOf(header);
    if (index < 0) {
      throw new Error(`Planning Table is missing required validation header: ${header}.`);
    }
    return index;
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

  function assertMainReportSpillReady(values, formulas, expectedFormula) {
    const anchorFormulaMatches = normalizeFormula(formulas[0][0]) === normalizeFormula(expectedFormula);
    if (anchorFormulaMatches && !isSpillError(values[0][0])) {
      return;
    }

    for (let rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
      for (let columnIndex = 0; columnIndex < values[rowIndex].length; columnIndex += 1) {
        if (anchorFormulaMatches && rowIndex === 0 && columnIndex === 0) {
          continue;
        }
        if (hasCellContent(values[rowIndex][columnIndex]) || hasCellContent(formulas[rowIndex][columnIndex])) {
          const address = `Planning Review!${columnName(columnIndex + 1)}${rowIndex + 4}`;
      throw new Error(`${address} blocks the main report spill. Clear Planning Review!A4:N200 or rerun Create Starter Sheets before inserting demo outputs.`);
        }
      }
    }
  }

  function hasCellContent(value) {
    return value !== null && value !== undefined && String(value).trim() !== "";
  }

  function isSpillError(value) {
    return String(value || "").trim().toUpperCase() === "#SPILL!";
  }

  function columnName(columnNumber) {
    let number = columnNumber;
    let name = "";
    while (number > 0) {
      const remainder = (number - 1) % 26;
      name = String.fromCharCode(65 + remainder) + name;
      number = Math.floor((number - 1) / 26);
    }
    return name;
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
      `- Dropdown lists ready: ${summary.dropdownListCount}`,
      `- Row validations configured: ${summary.rowValidationRuleCount}/${rowValidationRules.length}`
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
        formula: `=${compactFormulaBody(body)}`
      };
    });
  }

  function stripBlockComments(text) {
    return text.replace(/\/\*[\s\S]*?\*\//g, "");
  }

  function compactFormulaBody(text) {
    const source = stripBlockComments(text);
    let out = "";
    let inString = false;
    let inQuotedSheet = false;
    for (let i = 0; i < source.length; i++) {
      const ch = source[i];
      if (ch === '"' && !inQuotedSheet) {
        out += ch;
        if (inString && source[i + 1] === '"') {
          out += source[i + 1];
          i += 1;
        } else {
          inString = !inString;
        }
        continue;
      }
      if (ch === "'" && !inString) {
        out += ch;
        if (inQuotedSheet && source[i + 1] === "'") {
          out += source[i + 1];
          i += 1;
        } else {
          inQuotedSheet = !inQuotedSheet;
        }
        continue;
      }
      if (!inString && !inQuotedSheet && /\s/.test(ch)) {
        continue;
      }
      out += ch;
    }
    return out;
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
