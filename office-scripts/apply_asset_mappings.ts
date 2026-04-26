/**
 * Starter Office Script for applying accepted asset-mapping rows.
 *
 * This is workbook mechanics only. It does not export RDF or any external format.
 *
 * Expected target sheets/tables when present:
 * - Project Asset Map / tblProjectAssetMap
 * - Asset Changes / tblAssetChanges
 * - Asset State History / tblAssetStateHistory
 *
 * Source rows are read from the first available staging table candidate, commonly
 * tblSemanticAssets, tblAssetPromotionQueue, tblAssetMappingStaging, or tblProjectAssetMap.
 */
function main(workbook: ExcelScript.Workbook): {
    sourceTable: string;
    processed: number;
    applied: number;
    skipped: number;
    errors: number;
    timestampUtc: string;
    message: string;
} {
    type Cell = string | number | boolean | null;

    const timestamp = new Date().toISOString();

    const norm = (value: Cell | undefined): string => {
        if (value === null || value === undefined) return "";
        return String(value).trim();
    };

    const key = (value: string): string => value.toLowerCase().replace(/[\s_-]/g, "");

    const isTruthy = (value: Cell | undefined): boolean => {
        const text = norm(value).toLowerCase();
        return text === "true" || text === "yes" || text === "y" || text === "1";
    };

    const tableOrUndefined = (sheetName: string, tableNames: string[]): ExcelScript.Table | undefined => {
        const sheet = workbook.getWorksheet(sheetName);
        if (!sheet) return undefined;

        for (const tableName of tableNames) {
            try {
                return sheet.getTable(tableName);
            } catch (error) {
                // Keep looking. Workbooks may use only the sheet-level table in this starter phase.
            }
        }

        const tables = sheet.getTables();
        return tables.length > 0 ? tables[0] : undefined;
    };

    const sourceCandidates = [
        { sheet: "Semantic Assets", names: ["tblSemanticAssets", "tblAssetPromotionQueue", "tblAssetMappingStaging"] },
        { sheet: "Asset Setup", names: ["tblAssetPromotionQueue", "tblAssetMappingStaging"] },
        { sheet: "Project Asset Map", names: ["tblProjectAssetMap"] }
    ];

    let sourceTable: ExcelScript.Table | undefined;
    for (const candidate of sourceCandidates) {
        sourceTable = tableOrUndefined(candidate.sheet, candidate.names);
        if (sourceTable) break;
    }

    const mapTable = tableOrUndefined("Project Asset Map", ["tblProjectAssetMap"]);
    const changesTable = tableOrUndefined("Asset Changes", ["tblAssetChanges"]);
    const historyTable = tableOrUndefined("Asset State History", ["tblAssetStateHistory"]);
    const missingOptionalTargets: string[] = [];
    if (!changesTable) missingOptionalTargets.push("Asset Changes / tblAssetChanges");
    if (!historyTable) missingOptionalTargets.push("Asset State History / tblAssetStateHistory");

    if (!sourceTable) {
        return {
            sourceTable: "",
            processed: 0,
            applied: 0,
            skipped: 0,
            errors: 0,
            timestampUtc: timestamp,
            message: "No asset staging/source table found."
        };
    }

    if (!mapTable) {
        return {
            sourceTable: sourceTable.getName(),
            processed: 0,
            applied: 0,
            skipped: 0,
            errors: 0,
            timestampUtc: timestamp,
            message: "Project Asset Map table not found; nothing was applied."
        };
    }

    const headers = (table: ExcelScript.Table): string[] =>
        (table.getHeaderRowRange().getValues()[0] as Cell[]).map(value => norm(value));

    const sourceHeaders = headers(sourceTable);
    const mapHeaders = headers(mapTable);
    const changeHeaders = changesTable ? headers(changesTable) : [];
    const historyHeaders = historyTable ? headers(historyTable) : [];

    const indexAny = (headersList: string[], names: string[]): number => {
        const normalizedHeaders = headersList.map(header => key(header));
        for (const name of names) {
            const index = normalizedHeaders.indexOf(key(name));
            if (index >= 0) return index;
        }
        return -1;
    };

    const getAny = (headersList: string[], row: Cell[], names: string[]): string => {
        const index = indexAny(headersList, names);
        return index >= 0 ? norm(row[index]) : "";
    };

    const setAny = (headersList: string[], row: Cell[], names: string[], value: Cell): void => {
        const index = indexAny(headersList, names);
        if (index >= 0) row[index] = value;
    };

    const isReadyRow = (headersList: string[], row: Cell[]): boolean => {
        const applyStatus = getAny(headersList, row, ["ApplyStatus"]);
        if (applyStatus.toLowerCase() === "applied") return false;

        if (isTruthy(row[indexAny(headersList, ["ApplyReady", "Ready"])])) return true;

        const statusText = [
            getAny(headersList, row, ["PromotionStatus"]),
            getAny(headersList, row, ["MappingStatus"]),
            getAny(headersList, row, ["ApplyStatus"])
        ].join("|").toLowerCase();

        return (
            statusText.includes("accepted") ||
            statusText.includes("project_ready") ||
            statusText.includes("project-ready") ||
            statusText.includes("ready")
        );
    };

    const validate = (
        changeType: string,
        sourceAssetId: string,
        targetAssetId: string,
        installedState: string,
        evidenceId: string
    ): string[] => {
        const messages: string[] = [];
        if (changeType === "new_asset" && targetAssetId === "") {
            messages.push("new_asset requires target_asset_id");
        }
        if (changeType === "replace_asset" && (sourceAssetId === "" || targetAssetId === "")) {
            messages.push("replace_asset requires source_asset_id and target_asset_id");
        }
        if (changeType === "upgrade_asset" && sourceAssetId !== targetAssetId) {
            messages.push("upgrade_asset requires source_asset_id = target_asset_id");
        }
        if (installedState === "installed" && evidenceId === "") {
            messages.push("installed_state = installed requires evidence_id");
        }
        return messages;
    };

    const mapKey = (projectKey: string, assetId: string): string => `${key(projectKey)}|${key(assetId)}`;
    const sourceIsMapTable = sourceTable.getName() === mapTable.getName();
    const sourceRange = sourceTable.getRangeBetweenHeaderAndTotal();
    const sourceRows = sourceRange.getValues() as Cell[][];
    const mapRange = mapTable.getRangeBetweenHeaderAndTotal();
    const mapRows = sourceIsMapTable ? sourceRows : (mapRange.getValues() as Cell[][]);
    const originalSourceRowCount = sourceRows.length;

    const existingMapRows: { [key: string]: number } = {};
    for (let r = 0; r < mapRows.length; r++) {
        const projectKey = getAny(mapHeaders, mapRows[r], ["ProjectKey", "ProjectId"]);
        const assetId = getAny(mapHeaders, mapRows[r], ["AssetId", "TargetAssetId"]);
        if (projectKey && assetId) existingMapRows[mapKey(projectKey, assetId)] = r;
    }

    let processed = 0;
    let applied = 0;
    let skipped = 0;
    let errors = 0;
    let mapChanged = false;

    for (let r = 0; r < sourceRows.length; r++) {
        const row = sourceRows[r];
        if (!isReadyRow(sourceHeaders, row)) {
            skipped++;
            continue;
        }

        processed++;

        const projectKey = getAny(sourceHeaders, row, ["ProjectKey", "ProjectId"]);
        const projectDescription = getAny(sourceHeaders, row, ["ProjectDescription", "ProjectDesc", "ProjDesc"]);
        const assetLabel = getAny(sourceHeaders, row, ["AssetLabel", "AssetName"]);
        const assetType = getAny(sourceHeaders, row, ["AssetType"]);
        const changeType = getAny(sourceHeaders, row, ["ProposedChangeType", "ChangeType"]).toLowerCase();
        const sourceAssetId = getAny(sourceHeaders, row, ["SourceAssetId", "source_asset_id"]);
        const targetAssetId = getAny(sourceHeaders, row, ["TargetAssetId", "target_asset_id", "AssetId"]);
        const installedState = getAny(sourceHeaders, row, ["InstalledState", "installed_state", "AssetState"]).toLowerCase();
        const evidenceId = getAny(sourceHeaders, row, ["EvidenceId", "evidence_id"]);
        const effectiveAssetId = targetAssetId || sourceAssetId;

        const validationMessages = validate(changeType, sourceAssetId, targetAssetId, installedState, evidenceId);
        if (projectKey === "") validationMessages.push("ProjectKey is required");
        if (changeType === "") validationMessages.push("ChangeType is required");
        if (effectiveAssetId === "") validationMessages.push("AssetId is required");

        if (validationMessages.length > 0) {
            setAny(sourceHeaders, row, ["ApplyStatus"], "Error");
            setAny(sourceHeaders, row, ["AppliedOn"], timestamp);
            setAny(sourceHeaders, row, ["ApplyMessage"], validationMessages.join("; "));
            errors++;
            continue;
        }

        const existingIndex = existingMapRows[mapKey(projectKey, effectiveAssetId)];
        if (existingIndex === undefined) {
            const newMapRow = mapHeaders.map(() => "");
            setAny(mapHeaders, newMapRow, ["ProjectKey", "ProjectId"], projectKey);
            setAny(mapHeaders, newMapRow, ["ProjectDescription", "ProjectDesc"], projectDescription);
            setAny(mapHeaders, newMapRow, ["AssetId", "TargetAssetId"], effectiveAssetId);
            setAny(mapHeaders, newMapRow, ["AssetLabel", "AssetName"], assetLabel);
            setAny(mapHeaders, newMapRow, ["AssetType"], assetType);
            setAny(mapHeaders, newMapRow, ["AssetState", "InstalledState"], installedState || "mapped");
            setAny(mapHeaders, newMapRow, ["EvidenceId"], evidenceId);
            setAny(mapHeaders, newMapRow, ["MappingStatus"], "active");
            setAny(mapHeaders, newMapRow, ["ApplyStatus"], "Applied");
            setAny(mapHeaders, newMapRow, ["AppliedOn"], timestamp);
            setAny(mapHeaders, newMapRow, ["ApplyMessage"], "Inserted by apply_asset_mappings");
            mapTable.addRows(-1, [newMapRow]);
            existingMapRows[mapKey(projectKey, effectiveAssetId)] = mapRows.length;
            mapRows.push(newMapRow);
        } else {
            const mapRow = mapRows[existingIndex];
            setAny(mapHeaders, mapRow, ["ProjectDescription", "ProjectDesc"], projectDescription);
            setAny(mapHeaders, mapRow, ["AssetLabel", "AssetName"], assetLabel);
            setAny(mapHeaders, mapRow, ["AssetType"], assetType);
            setAny(mapHeaders, mapRow, ["AssetState", "InstalledState"], installedState || getAny(mapHeaders, mapRow, ["AssetState"]));
            setAny(mapHeaders, mapRow, ["EvidenceId"], evidenceId || getAny(mapHeaders, mapRow, ["EvidenceId"]));
            setAny(mapHeaders, mapRow, ["MappingStatus"], "active");
            setAny(mapHeaders, mapRow, ["ApplyStatus"], "Applied");
            setAny(mapHeaders, mapRow, ["AppliedOn"], timestamp);
            setAny(mapHeaders, mapRow, ["ApplyMessage"], "Updated by apply_asset_mappings");
            mapChanged = true;
        }

        if (changesTable) {
            const changeRow = changeHeaders.map(() => "");
            setAny(changeHeaders, changeRow, ["ChangeId"], `CHG-${r + 1}-${Date.now()}`);
            setAny(changeHeaders, changeRow, ["ProjectKey", "ProjectId"], projectKey);
            setAny(changeHeaders, changeRow, ["ChangeType", "ProposedChangeType"], changeType);
            setAny(changeHeaders, changeRow, ["SourceAssetId"], sourceAssetId);
            setAny(changeHeaders, changeRow, ["TargetAssetId", "AssetId"], effectiveAssetId);
            setAny(changeHeaders, changeRow, ["InstalledState", "AssetState"], installedState);
            setAny(changeHeaders, changeRow, ["EvidenceId"], evidenceId);
            setAny(changeHeaders, changeRow, ["ChangeStatus"], "applied");
            setAny(changeHeaders, changeRow, ["AppliedOn"], timestamp);
            setAny(changeHeaders, changeRow, ["ApplyMessage"], "Applied by apply_asset_mappings");
            changesTable.addRows(-1, [changeRow]);
        }

        if (historyTable) {
            const historyRow = historyHeaders.map(() => "");
            setAny(historyHeaders, historyRow, ["EventId"], `EVT-${r + 1}-${Date.now()}`);
            setAny(historyHeaders, historyRow, ["AssetId", "TargetAssetId"], effectiveAssetId);
            setAny(historyHeaders, historyRow, ["ProjectKey", "ProjectId"], projectKey);
            setAny(historyHeaders, historyRow, ["AssetState", "InstalledState"], installedState || "mapped");
            setAny(historyHeaders, historyRow, ["EvidenceId"], evidenceId);
            setAny(historyHeaders, historyRow, ["EventSource"], "apply_asset_mappings");
            setAny(historyHeaders, historyRow, ["EventOn", "AppliedOn"], timestamp);
            setAny(historyHeaders, historyRow, ["ApplyMessage"], `${changeType} applied`);
            historyTable.addRows(-1, [historyRow]);
        }

        setAny(sourceHeaders, row, ["ApplyStatus"], "Applied");
        setAny(sourceHeaders, row, ["AppliedOn"], timestamp);
        setAny(sourceHeaders, row, ["ApplyMessage"], `Applied ${changeType} for ${effectiveAssetId}`);
        applied++;
    }

    if (mapChanged && !sourceIsMapTable) {
        mapTable.getRangeBetweenHeaderAndTotal().setValues(mapRows);
    }
    if (sourceRows.length > 0) {
        const writeSourceRange = sourceRows.length === originalSourceRowCount
            ? sourceRange
            : sourceTable.getRangeBetweenHeaderAndTotal();
        writeSourceRange.setValues(sourceRows);
    }

    return {
        sourceTable: sourceTable.getName(),
        processed,
        applied,
        skipped,
        errors,
        timestampUtc: timestamp,
        message: "Asset mapping apply complete. " +
            (missingOptionalTargets.length > 0
                ? "Optional target tables missing: " + missingOptionalTargets.join(", ") + ". "
                : "") +
            "RDF/export was not run."
    };
}
