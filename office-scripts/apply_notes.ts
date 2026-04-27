/**
 * Apply planning updates from 'Decision Staging'!tblDecisionStaging into live 'Planning Table' range.
 * Budget headers are on row 2. Key header is 'Project Description'.
 * Updates: 'Planning Notes', 'Timeline', 'Comments', 'Status'
 * Writes back status fields to tblDecisionStaging: ApplyStatus, AppliedOn, ApplyMessage, BudgetRowFound
 * Run 1 reads Planning Review P:R, records ReviewRow, and refreshes formula-backed tblDecisionStaging rows.
 * Run 2 applies only rows already marked Prepared.
 * A later run with no Planning Review P:R inputs resets stale staging rows to one blank formula-backed row.
 * After a successful apply, clears the matching Planning Review P:R source inputs.
 * Updates Planning Review O1:R3 with the latest operator status after each normal run.
 */
function main(workbook: ExcelScript.Workbook): {
    phase: string;
    prepared: number;
    applied: number;
    skipped: number;
    errors: number;
    timestampUtc: string
} {
    // ---- Config ----
    workbook.refreshAllDataConnections();

    // Optional but safer for LET/XLOOKUP chains
    workbook.getApplication().calculate(ExcelScript.CalculationType.fullRebuild);

    const APPLY_SHEET_NAME = "Decision Staging";
    const APPLY_TABLE_NAME = "tblDecisionStaging";

    const BUDGET_SHEET_NAME = "Planning Table";
    const BUDGET_HEADER_ROW_1BASED = 2; // headers on row 2
    const REVIEW_SHEET_NAME = "Planning Review";
    const REVIEW_RANGE_ADDRESS = "A4:R200";
    const REVIEW_RANGE_ROW0 = 3; // A4:R200 starts on row 4
    const REVIEW_INPUT_COL0 = 15; // P:R source inputs
    const CONTROL_RANGE_ADDRESS = "O1:R3";

    const HDR_KEY = "Project Description";
    const HDR_PLANNING_NOTES = "Planning Notes";
    const HDR_TIMELINE = "Timeline";
    const HDR_COMMENTS = "Comments";
    const HDR_STATUS = "Status";
    const COMMENTS_ROW_HEIGHT_POINTS = 45;
    const STATUS_PREPARED = "Prepared";
    const STATUS_APPLIED = "Applied";
    const STATUS_BLOCKED = "Blocked";
    const STATUS_SKIPPED = "Skipped";
    const STATUS_ERROR = "Error";

    // ---- Helpers ----
    type TableCell = string | number | boolean | null;

    const norm = (v: TableCell | undefined): string => {
        if (v === undefined || v === null) return "";
        return String(v).trim();
    };

    const isTrue = (v: TableCell | undefined): boolean => {
        if (typeof v === "boolean") return v;
        const s = norm(v).toUpperCase();
        return s === "TRUE" || s === "1" || s === "YES";
    };

    const nowIso = (): string => new Date().toISOString(); // UTC timestamp

    type RunSummary = {
        phase: string;
        prepared: number;
        applied: number;
        skipped: number;
        errors: number;
        timestampUtc: string;
    };

    const writeApplyNotesControl = (
        phaseLabel: string,
        result: string,
        nextAction: string,
        timestampUtc: string
    ): void => {
        const controlSheet = workbook.getWorksheet(REVIEW_SHEET_NAME);
        if (!controlSheet) return;

        const displayTimestamp = timestampUtc.substring(0, 19).replace("T", " ") + " UTC";
        const controlRange = controlSheet.getRange(CONTROL_RANGE_ADDRESS);
        controlRange.setValues([
            ["ApplyNotes Control", "Last Run", "Result", "Next Action"],
            [phaseLabel, displayTimestamp, result, nextAction],
            ["Source: Planning Review P:R", "Staging: Decision Staging", "Writeback: Planning Table", "Column O refreshes after apply"]
        ]);
        controlRange.getFormat().setWrapText(true);
        controlRange.getFormat().getFill().setColor("#EAF4E2");
        controlSheet.getRange("O1:R1").getFormat().getFont().setBold(true);
    };

    const finish = (
        phase: string,
        prepared: number,
        applied: number,
        skipped: number,
        errors: number,
        timestampUtc: string,
        result: string,
        nextAction: string
    ): RunSummary => {
        writeApplyNotesControl(phase, result, nextAction, timestampUtc);
        return {
            phase,
            prepared,
            applied,
            skipped,
            errors,
            timestampUtc
        };
    };

    // ---- Get Apply table ----
    const applySheet = workbook.getWorksheet(APPLY_SHEET_NAME);
    if (!applySheet) throw new Error(`Worksheet not found: ${APPLY_SHEET_NAME}`);

    const applyTable = applySheet.getTable(APPLY_TABLE_NAME);
    if (!applyTable) throw new Error(`Table not found on '${APPLY_SHEET_NAME}': ${APPLY_TABLE_NAME}`);

    const applyRange = applyTable.getRange();
    const applyValues = applyRange.getValues() as TableCell[][];

    if (applyValues.length < 2) {
        return finish(
            "idle",
            0,
            0,
            0,
            0,
            nowIso(),
            "No staging rows are available.",
            "Run Setup Notes Workflow, then type updates in Planning Review P:R."
        );
    }

    const applyHeaders = applyValues[0].map(h => norm(h));
    const data = applyValues.slice(1);

    const colIndex = (name: string): number => {
        for (let i = 0; i < applyHeaders.length; i++) {
            if (applyHeaders[i] === name) return i;
        }
        throw new Error(`Missing column in ${APPLY_TABLE_NAME}: '${name}'`);
    };

    // Required input columns
    const iReviewRow = colIndex("ReviewRow");
    const iProjDesc = colIndex("ProjDesc");
    const iApplyReady = colIndex("ApplyReady");
    const iKeyStatus = colIndex("KeyStatus");
    const iMatchCount = colIndex("BudgetMatchCount");
    const iApplyAction = colIndex("ApplyAction");

    const iPlanningNotesNew = colIndex("PlanningNotes_New");
    const iTimelineNew = colIndex("Timeline_New");
    const iCommentsNew = colIndex("Comments_New");
    const iNewStatus = colIndex("NewStatus");
    const iStatusNew = colIndex("Status_New");

    // ---- Get Budget sheet + header mapping ----
    const budgetSheet = workbook.getWorksheet(BUDGET_SHEET_NAME);
    if (!budgetSheet) throw new Error(`Worksheet not found: ${BUDGET_SHEET_NAME}`);

    const used = budgetSheet.getUsedRange();
    if (!used) throw new Error(`No used range found on sheet: ${BUDGET_SHEET_NAME}`);

    const usedValues = used.getValues() as ExcelScript.RangeValueType[][];
    const usedRowCount = used.getRowCount();
    const usedRow0 = used.getRowIndex();    // 0-based
    const usedCol0 = used.getColumnIndex(); // 0-based

    const headerRow0 = BUDGET_HEADER_ROW_1BASED - 1; // 0-based absolute
    if (headerRow0 < usedRow0 || headerRow0 >= usedRow0 + usedRowCount) {
        throw new Error(`Budget header row ${BUDGET_HEADER_ROW_1BASED} is outside the used range.`);
    }

    const headerOffset = headerRow0 - usedRow0;
    const headerRowValues = usedValues[headerOffset].map(h => norm(h as TableCell));

    const budgetColByHeader = (hdrName: string): number => {
        for (let i = 0; i < headerRowValues.length; i++) {
            if (headerRowValues[i] === hdrName) return usedCol0 + i; // absolute col index
        }
        throw new Error(`Header '${hdrName}' not found on '${BUDGET_SHEET_NAME}' row ${BUDGET_HEADER_ROW_1BASED}`);
    };

    const cKey = budgetColByHeader(HDR_KEY);
    const cPlanningNotes = budgetColByHeader(HDR_PLANNING_NOTES);
    const cTL = budgetColByHeader(HDR_TIMELINE);
    const cCM = budgetColByHeader(HDR_COMMENTS);
    const cST = budgetColByHeader(HDR_STATUS);

    // Determine budget data rows
    const budgetDataStartRow0 = headerRow0 + 1; // row 3 if header is row 2
    const budgetLastRow0 = usedRow0 + usedRowCount - 1;
    if (budgetDataStartRow0 > budgetLastRow0) throw new Error("Budget has no data rows beneath the header.");

    const budgetDataRowCount = budgetLastRow0 - budgetDataStartRow0 + 1;

    const pmRange = budgetSheet.getRangeByIndexes(budgetDataStartRow0, cPlanningNotes, budgetDataRowCount, 1);
    const tlRange = budgetSheet.getRangeByIndexes(budgetDataStartRow0, cTL, budgetDataRowCount, 1);
    const cmRange = budgetSheet.getRangeByIndexes(budgetDataStartRow0, cCM, budgetDataRowCount, 1);
    const stRange = budgetSheet.getRangeByIndexes(budgetDataStartRow0, cST, budgetDataRowCount, 1);

    const pmVals = pmRange.getValues() as TableCell[][];
    const tlVals = tlRange.getValues() as TableCell[][];
    const cmVals = cmRange.getValues() as TableCell[][];
    const stVals = stRange.getValues() as TableCell[][];

    // ---- Apply table output/input ranges ----
    const applyStatusRange = applyTable.getColumnByName("ApplyStatus").getRangeBetweenHeaderAndTotal();
    const appliedOnRange = applyTable.getColumnByName("AppliedOn").getRangeBetweenHeaderAndTotal();
    const foundRange = applyTable.getColumnByName("BudgetRowFound").getRangeBetweenHeaderAndTotal();
    const msgRange = applyTable.getColumnByName("ApplyMessage").getRangeBetweenHeaderAndTotal();

    const newNoteRange = applyTable.getColumnByName("NewPlanningNotes").getRangeBetweenHeaderAndTotal();
    const newTLRange = applyTable.getColumnByName("NewTimeline").getRangeBetweenHeaderAndTotal();
    const newStatusRange = applyTable.getColumnByName("NewStatus").getRangeBetweenHeaderAndTotal();

    const applyStatusVals = applyStatusRange.getValues() as TableCell[][];
    const appliedOnVals = appliedOnRange.getValues() as TableCell[][];
    const foundVals = foundRange.getValues() as TableCell[][];
    const msgVals = msgRange.getValues() as TableCell[][];
    const newNoteVals = newNoteRange.getValues() as TableCell[][];
    const newTLVals = newTLRange.getValues() as TableCell[][];
    const newStatusVals = newStatusRange.getValues() as TableCell[][];

    // ---- Planning Review source inputs ----
    const reviewSheet = workbook.getWorksheet(REVIEW_SHEET_NAME);
    const reviewRange = reviewSheet ? reviewSheet.getRange(REVIEW_RANGE_ADDRESS) : undefined;
    const reviewValues = reviewRange ? reviewRange.getValues() as TableCell[][] : [];
    const reviewHeaders = reviewValues.length > 0 ? reviewValues[0].map(h => norm(h).split("\n")[0]) : [];
    const reviewProjDescIndex = reviewHeaders.indexOf(HDR_KEY);

    const findReviewInputRowOffset = (
        reviewRow1: number | undefined,
        projDesc: string,
        rawNote: string,
        rawTimeline: string,
        rawStatus: string
    ): number | undefined => {
        if (!reviewSheet || reviewProjDescIndex < 0 || reviewValues.length < 2) return undefined;

        if (reviewRow1 !== undefined && reviewRow1 > REVIEW_RANGE_ROW0 + 1) {
            const offset = reviewRow1 - REVIEW_RANGE_ROW0 - 1;
            if (offset > 0 && offset < reviewValues.length) {
                const reviewRow = reviewValues[offset];
                if (
                    norm(reviewRow[reviewProjDescIndex]) === projDesc &&
                    norm(reviewRow[15]) === rawNote &&
                    norm(reviewRow[16]) === rawTimeline &&
                    norm(reviewRow[17]) === rawStatus
                ) {
                    return offset;
                }
            }
        }

        for (let r = 1; r < reviewValues.length; r++) {
            if (norm(reviewValues[r][reviewProjDescIndex]) !== projDesc) continue;
            if (
                norm(reviewValues[r][15]) === rawNote &&
                norm(reviewValues[r][16]) === rawTimeline &&
                norm(reviewValues[r][17]) === rawStatus
            ) {
                return r;
            }
        }

        return undefined;
    };

    const clearReviewInputs = (rowOffset: number): void => {
        if (!reviewSheet) return;
        reviewSheet.getRangeByIndexes(REVIEW_RANGE_ROW0 + rowOffset, REVIEW_INPUT_COL0, 1, 3).setValues([["", "", ""]]);
    };

    // Pull key column values once
    const keyColRange = budgetSheet.getRangeByIndexes(budgetDataStartRow0, cKey, budgetDataRowCount, 1);
    const keyColVals = keyColRange.getValues() as ExcelScript.RangeValueType[][];
    const keys: string[] = [];
    for (let i = 0; i < keyColVals.length; i++) {
        keys.push(norm(keyColVals[i]?.[0] as TableCell));
    }

    // Build maps using plain objects
    const keyToRow: { [k: string]: number } = {};
    const dupKey: { [k: string]: boolean } = {};
    const keyCount: { [k: string]: number } = {};

    for (let i = 0; i < keys.length; i++) {
        const k = keys[i];
        if (!k) continue;
        keyCount[k] = (keyCount[k] || 0) + 1;
        const absRow1 = (budgetDataStartRow0 + i) + 1; // 1-based row number
        if (keyToRow[k] !== undefined) {
            dupKey[k] = true;
        } else {
            keyToRow[k] = absRow1;
        }
    }

    // Remove duplicates from keyToRow
    for (const k in dupKey) {
        delete keyToRow[k];
    }

    const timestamp = nowIso();
    const app = workbook.getApplication();

    const flushApplyTable = (): void => {
        applyStatusRange.setValues(applyStatusVals);
        appliedOnRange.setValues(appliedOnVals);
        foundRange.setValues(foundVals);
        msgRange.setValues(msgVals);
    };

    const blockMessage = (
        projDesc: string,
        applyReady: boolean,
        keyStatus: string,
        matchCount: number,
        budgetRow1: number | undefined
    ): string => {
        if (!projDesc) return "Blocked: Project Description is blank. Enter notes beside a populated Planning Review report row.";
        if (matchCount !== 1) return `Blocked: expected exactly 1 Planning Table match for '${projDesc}', found ${isNaN(matchCount) ? "blank" : String(matchCount)}.`;
        if (dupKey[projDesc]) return `Blocked: Planning Table has duplicate rows for '${projDesc}'.`;
        if (budgetRow1 === undefined) return `Blocked: Planning Table row for '${projDesc}' could not be resolved.`;
        if (keyStatus !== "OK") return `Blocked: KeyStatus is '${keyStatus || "blank"}'.`;
        if (!applyReady) return "Blocked: ApplyReady is not TRUE.";
        return "Blocked: row is not eligible to apply.";
    };

    const duplicateTargetMessage = (budgetRow1: number, count: number): string => {
        return `Blocked: ${count} Planning Review rows target Planning Table row ${budgetRow1}. Leave one update per target row, then run ApplyNotes again.`;
    };

    const isApplyEligible = (
        projDesc: string,
        applyReady: boolean,
        keyStatus: string,
        matchCount: number,
        budgetRow1: number | undefined
    ): boolean => {
        return projDesc !== ""
            && applyReady
            && keyStatus === "OK"
            && matchCount === 1
            && !dupKey[projDesc]
            && budgetRow1 !== undefined;
    };

    const prepMessage = (
        projDesc: string,
        applyReady: boolean,
        keyStatus: string,
        matchCount: number,
        budgetRow1: number | undefined
    ): string => {
        if (!isApplyEligible(projDesc, applyReady, keyStatus, matchCount, budgetRow1)) {
            return blockMessage(projDesc, applyReady, keyStatus, matchCount, budgetRow1);
        }
        return `Prepared: matched Planning Table row ${budgetRow1}. Review Decision Staging, then run ApplyNotes again to apply.`;
    };

    const hasPreparedStagingRows = (): boolean => {
        for (let r = 0; r < data.length; r++) {
            const rawNote = norm(newNoteVals[r]?.[0]);
            const rawTL = norm(newTLVals[r]?.[0]);
            const rawStatus = norm(newStatusVals[r]?.[0]);
            const hasRawInput = rawNote !== "" || rawTL !== "" || rawStatus !== "";
            if (hasRawInput && norm(applyStatusVals[r]?.[0]) === STATUS_PREPARED) return true;
        }

        return false;
    };

    const reviewHeaderIndex = (header: string): number => {
        for (let i = 0; i < reviewHeaders.length; i++) {
            if (reviewHeaders[i] === header) return i;
        }
        throw new Error(`Missing Planning Review header: '${header}'`);
    };

    type PrepareRow = {
        reviewRow: TableCell;
        applyStatus: string;
        budgetRowFound: TableCell;
        applyMessage: string;
    };

    type ReviewCandidate = {
        reviewRow: number;
        projDesc: string;
        matchCount: number;
        budgetRow1: number | undefined;
        keyStatus: string;
        applyReady: boolean;
    };

    const buildReviewPrepareRows = (): PrepareRow[] => {
        if (!reviewSheet || reviewValues.length < 2) return [];

        const iDesc = reviewHeaderIndex(HDR_KEY);
        const candidates: ReviewCandidate[] = [];

        for (let r = 1; r < reviewValues.length; r++) {
            const reviewRow = reviewValues[r];
            const rawNote = norm(reviewRow[15]);
            const rawTimeline = norm(reviewRow[16]);
            const rawStatus = norm(reviewRow[17]);
            const hasRawInput = rawNote !== "" || rawTimeline !== "" || rawStatus !== "";
            if (!hasRawInput) continue;

            const projDesc = norm(reviewRow[iDesc]);
            const matchCount = projDesc === "" ? 0 : (keyCount[projDesc] || 0);
            const budgetRow1 = matchCount === 1 ? keyToRow[projDesc] : undefined;
            const keyStatus = projDesc === "" ? "" : (matchCount === 1 ? "OK" : "BLOCKED");
            const applyReady = projDesc !== "" && matchCount === 1;

            candidates.push({
                reviewRow: REVIEW_RANGE_ROW0 + r + 1,
                projDesc,
                matchCount,
                budgetRow1,
                keyStatus,
                applyReady
            });
        }

        const targetCounts: { [row: string]: number } = {};
        for (const candidate of candidates) {
            if (candidate.budgetRow1 === undefined) continue;
            const rowKey = String(candidate.budgetRow1);
            targetCounts[rowKey] = (targetCounts[rowKey] || 0) + 1;
        }

        const rows: PrepareRow[] = [];
        for (const candidate of candidates) {
            const duplicateTargetCount = candidate.budgetRow1 === undefined
                ? 0
                : (targetCounts[String(candidate.budgetRow1)] || 0);
            const hasDuplicateTarget = candidate.budgetRow1 !== undefined && duplicateTargetCount > 1;
            const eligible = !hasDuplicateTarget
                && isApplyEligible(candidate.projDesc, candidate.applyReady, candidate.keyStatus, candidate.matchCount, candidate.budgetRow1);

            rows.push({
                reviewRow: candidate.reviewRow,
                applyStatus: eligible ? STATUS_PREPARED : STATUS_BLOCKED,
                budgetRowFound: candidate.budgetRow1 === undefined ? "" : candidate.budgetRow1,
                applyMessage: hasDuplicateTarget && candidate.budgetRow1 !== undefined
                    ? duplicateTargetMessage(candidate.budgetRow1, duplicateTargetCount)
                    : prepMessage(candidate.projDesc, candidate.applyReady, candidate.keyStatus, candidate.matchCount, candidate.budgetRow1)
            });
        }

        return rows;
    };

    const setColumnFormulas = (columnName: string, formulas: string[]): void => {
        applyTable.getColumnByName(columnName)
            .getRangeBetweenHeaderAndTotal()
            .setFormulas(formulas.map(formula => [formula]));
    };

    const setColumnValues = (columnName: string, values: TableCell[]): void => {
        applyTable.getColumnByName(columnName)
            .getRangeBetweenHeaderAndTotal()
            .setValues(values.map(value => [value]));
    };

    const repeatedFormulas = (rowCount: number, formula: string): string[] => {
        const formulas: string[] = [];
        for (let i = 0; i < rowCount; i++) formulas.push(formula);
        return formulas;
    };

    const indexedNotesFormula = (sourceIndex: number): string => {
        const sourceRows = "DROP(Notes.FromArrayv,1)";
        return `=IF([@ReviewRow]="","",IFERROR(INDEX(${sourceRows},XMATCH([@ReviewRow],CHOOSECOLS(${sourceRows},1),0),${sourceIndex}),""))`;
    };

    const refreshFormulaBackedApplyTableRows = (rows: PrepareRow[]): void => {
        const rowCount = rows.length;
        const currentRowCount = applyTable.getRangeBetweenHeaderAndTotal().getRowCount();

        if (currentRowCount < rowCount) {
            const blankRows: TableCell[][] = [];
            for (let r = currentRowCount; r < rowCount; r++) {
                blankRows.push(applyHeaders.map(() => ""));
            }
            applyTable.addRows(-1, blankRows);
        } else if (currentRowCount > rowCount) {
            applyTable.deleteRowsAt(rowCount, currentRowCount - rowCount);
        }

        applyTable.getRangeBetweenHeaderAndTotal().clear(ExcelScript.ClearApplyTo.contents);

        const stagingColumns: [string, number][] = [
            ["GroupType", 2],
            ["GroupValue", 3],
            ["Category", 4],
            ["ProjDesc", 5],
            ["AnnualProj", 6],
            ["ActualsYTD", 7],
            ["ExistingMeetingNotes", 8],
            ["NewPlanningNotes", 9],
            ["NewTimeline", 10],
            ["NewStatus", 11]
        ];
        for (const [columnName, sourceIndex] of stagingColumns) {
            setColumnFormulas(columnName, repeatedFormulas(rowCount, indexedNotesFormula(sourceIndex)));
        }

        setColumnFormulas("ApplyAction", repeatedFormulas(rowCount, '=IF(OR([@NewPlanningNotes]<>"",[@NewTimeline]<>"",[@NewStatus]<>""),"NOTE_TIMELINE_STATUS","")'));
        setColumnFormulas("PlanningNotes_New", repeatedFormulas(rowCount, '=IF([@NewPlanningNotes]<>"",[@NewPlanningNotes],"")'));
        setColumnFormulas("Timeline_New", repeatedFormulas(rowCount, '=IF([@NewTimeline]<>"",[@NewTimeline],"")'));
        setColumnFormulas("Comments_New", repeatedFormulas(rowCount, '=""'));
        setColumnFormulas("Status_New", repeatedFormulas(rowCount, '=IF([@NewStatus]<>"",[@NewStatus],"")'));
        setColumnFormulas("BudgetMatchCount", repeatedFormulas(rowCount, '=IF([@ProjDesc]="","",SUMPRODUCT(--(INDEX(\'Planning Table\'!$A$3:$BM$200,,XMATCH("Project Description",\'Planning Table\'!$A$2:$BM$2,0))=[@ProjDesc])))'));
        setColumnFormulas("KeyStatus", repeatedFormulas(rowCount, '=IF([@ProjDesc]="","",IF([@BudgetMatchCount]=1,"OK","BLOCKED"))'));
        setColumnFormulas("ApplyReady", repeatedFormulas(rowCount, '=AND([@ProjDesc]<>"",[@BudgetMatchCount]=1,OR([@NewPlanningNotes]<>"",[@NewTimeline]<>"",[@NewStatus]<>""))'));

        setColumnValues("ReviewRow", rows.map(row => row.reviewRow));
        setColumnValues("ApplyStatus", rows.map(row => row.applyStatus));
        setColumnValues("AppliedOn", rows.map(() => ""));
        setColumnValues("ApplyMessage", rows.map(row => row.applyMessage));
        setColumnValues("BudgetRowFound", rows.map(row => row.budgetRowFound));
    };

    const resetFormulaBackedApplyTable = (): void => {
        refreshFormulaBackedApplyTableRows([{
            reviewRow: "",
            applyStatus: "",
            budgetRowFound: "",
            applyMessage: ""
        }]);
    };

    const needsStagingReset = (): boolean => {
        if (data.length !== 1) return true;
        for (let r = 0; r < data.length; r++) {
            if (norm(data[r]?.[iReviewRow]) !== "") return true;
            if (norm(applyStatusVals[r]?.[0]) !== "") return true;
            if (norm(appliedOnVals[r]?.[0]) !== "") return true;
            if (norm(foundVals[r]?.[0]) !== "") return true;
            if (norm(msgVals[r]?.[0]) !== "") return true;
        }
        return false;
    };

    const preparedRowsPending = hasPreparedStagingRows();

    if (!preparedRowsPending) {
        const reviewPrepareRows = buildReviewPrepareRows();
        if (reviewPrepareRows.length > 0) {
            refreshFormulaBackedApplyTableRows(reviewPrepareRows);
            app.calculate(ExcelScript.CalculationType.fullRebuild);
            const preparedCount = reviewPrepareRows.filter(row => row.applyStatus === STATUS_PREPARED).length;
            const blockedCount = reviewPrepareRows.length - preparedCount;
            return finish(
                "prepare",
                preparedCount,
                0,
                blockedCount,
                0,
                timestamp,
                `${preparedCount} row(s) prepared; ${blockedCount} blocked.`,
                preparedCount > 0
                    ? "Review Decision Staging, then run ApplyNotes again."
                    : "Fix blocked rows or enter valid Planning Review P:R updates."
            );
        }

        if (needsStagingReset()) {
            resetFormulaBackedApplyTable();
            app.calculate(ExcelScript.CalculationType.fullRebuild);
            return finish(
                "reset",
                0,
                0,
                0,
                0,
                timestamp,
                "No current P:R inputs. Decision Staging reset.",
                "Type updates in Planning Review P:R, then run ApplyNotes."
            );
        }
    }

    // ---- Apply updates ----
    let preparedRemaining = 0;
    let appliedCount = 0;
    let skippedCount = 0;
    let errorCount = 0;
    let stagedCount = 0;
    const commentsRowsToFormat: number[] = [];

    const markCommentsRowForFormat = (budgetRow0: number): void => {
        for (const row0 of commentsRowsToFormat) {
            if (row0 === budgetRow0) return;
        }
        commentsRowsToFormat.push(budgetRow0);
    };

    const preparedTargetCounts: { [row: string]: number } = {};
    for (let r = 0; r < data.length; r++) {
        const rawNote = norm(newNoteVals[r]?.[0]);
        const rawTL = norm(newTLVals[r]?.[0]);
        const rawStatus = norm(newStatusVals[r]?.[0]);
        const hasRawInput = rawNote !== "" || rawTL !== "" || rawStatus !== "";
        if (!hasRawInput || norm(applyStatusVals[r]?.[0]) !== STATUS_PREPARED) continue;

        const projDesc = norm(data[r][iProjDesc]);
        const budgetRow1 = projDesc && !dupKey[projDesc] ? keyToRow[projDesc] : undefined;
        if (budgetRow1 === undefined) continue;

        const rowKey = String(budgetRow1);
        preparedTargetCounts[rowKey] = (preparedTargetCounts[rowKey] || 0) + 1;
    }

    for (let r = 0; r < data.length; r++) {
        const row = data[r];
        const rawNote = norm(newNoteVals[r][0]);
        const rawTL = norm(newTLVals[r][0]);
        const rawStatus = norm(newStatusVals[r][0]);
        const hasRawInput = rawNote !== "" || rawTL !== "" || rawStatus !== "";
        const applyStatus = norm(applyStatusVals[r][0]);

        if (!hasRawInput || applyStatus !== STATUS_PREPARED) continue;

        stagedCount++;

        const projDesc = norm(row[iProjDesc]);
        const applyReady = isTrue(row[iApplyReady]);
        const keyStatus = norm(row[iKeyStatus]);
        const matchCountRaw = norm(row[iMatchCount]);
        const matchCount = matchCountRaw === "" ? NaN : Number(matchCountRaw);
        const action = norm(row[iApplyAction]).toUpperCase();
        const budgetRow1 = projDesc && !dupKey[projDesc] ? keyToRow[projDesc] : undefined;
        const reviewRowRaw = Number(norm(row[iReviewRow]));
        const reviewRow1 = isNaN(reviewRowRaw) ? undefined : reviewRowRaw;
        const duplicatePreparedTargetCount = budgetRow1 === undefined
            ? 0
            : (preparedTargetCounts[String(budgetRow1)] || 0);

        if (budgetRow1 !== undefined && duplicatePreparedTargetCount > 1) {
            applyStatusVals[r][0] = STATUS_BLOCKED;
            appliedOnVals[r][0] = "";
            foundVals[r][0] = budgetRow1;
            msgVals[r][0] = duplicateTargetMessage(budgetRow1, duplicatePreparedTargetCount);
            skippedCount++;
            continue;
        }

        if (!isApplyEligible(projDesc, applyReady, keyStatus, matchCount, budgetRow1)) {
            applyStatusVals[r][0] = STATUS_BLOCKED;
            appliedOnVals[r][0] = "";
            foundVals[r][0] = budgetRow1 === undefined ? "" : budgetRow1;
            msgVals[r][0] = blockMessage(projDesc, applyReady, keyStatus, matchCount, budgetRow1);
            skippedCount++;
            continue;
        }

        const writeNote = action.indexOf("NOTE") >= 0;
        const writeTL = action.indexOf("TIMELINE") >= 0;
        const newStatusRaw = norm(row[iNewStatus]);
        const writeStatus = action.indexOf("STATUS") >= 0 || newStatusRaw !== "";

        const pmNew = norm(row[iPlanningNotesNew]);
        const tlNew = norm(row[iTimelineNew]);
        const cmNew = norm(row[iCommentsNew]);
        const statusNew = norm(row[iStatusNew]);

        const budgetRow0 = budgetRow1 - 1; // 0-based
        const bufRow = budgetRow0 - budgetDataStartRow0;
        const messages: string[] = [];
        let commentsTouched = false;

        try {
            const pmOld = norm(pmVals[bufRow][0]);
            const cmOld = norm(cmVals[bufRow][0]);

            if (writeNote && pmNew !== "") {
                // 1. Archive old Planning Notes -> Comments
                if (pmOld !== "") {
                    const archiveLine = timestamp.substring(0, 10) + " | " + pmOld;
                    cmVals[bufRow][0] = archiveLine + (cmOld !== "" ? ("\n" + cmOld) : "");
                    commentsTouched = true;
                }

                // 2. Write new Planning Notes clean (no timestamp)
                pmVals[bufRow][0] = pmNew.replace(/^\d{4}-\d{2}-\d{2}\s*\|\s*/, "");
                messages.push(pmOld !== "" ? "Planning Notes (prior note archived)" : "Planning Notes");
            }

            if (writeTL && tlNew !== "") {
                tlVals[bufRow][0] = tlNew;
                messages.push("Timeline");
            }

            if (cmNew !== "") {
                cmVals[bufRow][0] = cmNew;
                messages.push("Comments");
                commentsTouched = true;
            }

            if (writeStatus && statusNew !== "") {
                stVals[bufRow][0] = statusNew;
                messages.push("Status");
            }

            if (messages.length === 0) {
                applyStatusVals[r][0] = STATUS_SKIPPED;
                appliedOnVals[r][0] = "";
                foundVals[r][0] = budgetRow1;
                msgVals[r][0] = "Skipped: no non-empty target values were available to write.";
                skippedCount++;
                continue;
            }

            applyStatusVals[r][0] = STATUS_APPLIED;
            appliedOnVals[r][0] = timestamp;
            foundVals[r][0] = budgetRow1;
            msgVals[r][0] = "Applied: updated " + messages.join(" + ");

            appliedCount++;
            if (commentsTouched) {
                markCommentsRowForFormat(budgetRow0);
            }

            const reviewRowOffset = findReviewInputRowOffset(reviewRow1, projDesc, rawNote, rawTL, rawStatus);
            if (reviewRowOffset !== undefined) {
                clearReviewInputs(reviewRowOffset);
                msgVals[r][0] = norm(msgVals[r][0]) + "; cleared Planning Review P:R";
            } else {
                msgVals[r][0] = norm(msgVals[r][0]) + "; Planning Review P:R was not cleared because the source row no longer matched";
            }
        } catch (e: unknown) {
            const msg = (e instanceof Error) ? e.message : String(e);
            applyStatusVals[r][0] = STATUS_ERROR;
            appliedOnVals[r][0] = timestamp;
            foundVals[r][0] = budgetRow1;
            msgVals[r][0] = "Error: " + msg;
            errorCount++;
        }
    }

    if (stagedCount === 0) {
        return finish(
            "idle",
            0,
            0,
            0,
            0,
            timestamp,
            "No prepared staging rows were available to apply.",
            "Type updates in P:R or run once to prepare rows."
        );
    }

    // ---- Write back to Budget ----
    pmRange.setValues(pmVals);
    tlRange.setValues(tlVals);
    cmRange.setValues(cmVals);
    stRange.setValues(stVals);

    // ---- Write back to Apply table output buffers ----
    flushApplyTable();

    // Force a visible settle pass after source review inputs are cleared.
    app.calculate(ExcelScript.CalculationType.fullRebuild);

    // Keep archived comments readable without letting row height grow indefinitely.
    for (const row0 of commentsRowsToFormat) {
        budgetSheet.getRangeByIndexes(row0, cCM, 1, 1).getFormat().setWrapText(true);
        budgetSheet.getRangeByIndexes(row0, usedCol0, 1, used.getColumnCount()).getFormat().setRowHeight(COMMENTS_ROW_HEIGHT_POINTS);
    }

    return finish(
        "apply",
        preparedRemaining,
        appliedCount,
        skippedCount,
        errorCount,
        timestamp,
        `${appliedCount} row(s) applied; ${skippedCount} skipped or blocked; ${errorCount} error(s).`,
        errorCount > 0 || skippedCount > 0
            ? "Review Decision Staging ApplyMessage before running again."
            : "Review Planning Table. Matching P:R inputs were cleared."
    );
}
