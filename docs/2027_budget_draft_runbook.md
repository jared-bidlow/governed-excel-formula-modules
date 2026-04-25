# 2027 Budget Draft Runbook

## Purpose

Use this runbook to draft a 2027 monthly spend plan that fits the finance-provided monthly envelope while keeping detail by region, site, and job.

This process does not change the 2026 CapEx tracker contract. It uses the existing `Analysis` module as the starting view for current jobs, then builds a separate 2027 planning table.

## Inputs

You need four inputs before drafting:

1. Finance monthly cap for 2027.
2. Existing-job detail from `Analysis.PM_SPEND_REPORT()` and `Analysis.PM_SPEND_REPORT("Region")`.
3. New-job requests for 2027.
4. Optional pressure-test signals from `Ready` and `defer`.

## Working Tables

Maintain three planning tables.

### 1. FinanceMonthlyCap

One row per month.

Recommended columns:

- `BudgetYear`
- `MonthNum`
- `MonthName`
- `FinanceCap`

### 2. JobMaster2027

One row per job.

Recommended columns:

- `JobID`
- `JobDescription`
- `Region`
- `Site`
- `PM`
- `Category`
- `JobType`
- `Source`
- `Status`
- `Priority`
- `PlanningBucket`
- `DraftStartMonth`
- `PhaseTemplate`
- `IncludeInPlan`
- `2026Projected`
- `2026YTDSpend`
- `Remaining2026`
- `2027AnnualAsk`
- `Notes`

Use `JobType = Existing` for rows carried from the current portfolio and `JobType = New` for newly requested work.

Use `Source = Analysis` for existing jobs loaded from the current tracker and `Source = Intake` for new jobs entered directly.

### 3. JobMonthlyPlan2027

One row per `job x month`.

Recommended columns:

- `BudgetYear`
- `MonthNum`
- `MonthName`
- `JobID`
- `Region`
- `Site`
- `PM`
- `Category`
- `JobDescription`
- `JobType`
- `PlanningBucket`
- `DraftStartMonth`
- `PhaseTemplate`
- `IncludeInPlan`
- `AnnualAsk`
- `FinanceCapMonth`
- `ManualProjectedSpend`
- `RawWeight`
- `NormalizedWeight`
- `ProjectedSpend`

This is the main planning table. All rollups should come from this table.

Only `JobID` and `MonthNum` need to be seeded as the row drivers. The remaining columns can be formula-driven from `JobMaster2027` and `FinanceMonthlyCap`.

## Workflow

### Step 1. Load the monthly envelope

Enter the finance-provided amount for each month into `FinanceMonthlyCap`.

Do not skip months. Zero-cap months should still be listed explicitly.

### Step 2. Seed existing jobs

Use:

- `Analysis.PM_SPEND_REPORT()`
- `Analysis.PM_SPEND_REPORT("Region")`

Pull the job rows, not just the PM summary rows.

For each existing job, carry these fields into `JobMaster2027`:

- `Region`
- `Site`
- `PM`
- `Category`
- `Project Description`
- `2026 Projected`
- `YTD Spend`
- `Remaining`
- `Status`

Then assign:

- `JobType = Existing`
- `Source = Analysis`

### Step 3. Add new jobs

Enter new 2027 requests into `JobMaster2027` with:

- region,
- site,
- PM,
- category,
- description,
- proposed annual ask,
- priority,
- draft start month,
- notes.

Set:

- `JobType = New`
- `Source = Intake`

### Step 4. Pressure-test the list

Before phasing money into months, review the candidate jobs:

- Use `Ready` signals to separate execution-ready jobs from immature jobs.
- Use `defer` signals to identify jobs that should not automatically consume early-year capacity.

Typical screening buckets:

- `Committed carryover`
- `Candidate carryover`
- `Ready new ask`
- `Immature new ask`
- `Defer / hold`

### Step 5. Set the annual ask

Do not assume `Remaining` from the analysis view is automatically the 2027 ask.

For each job, choose an explicit `2027AnnualAsk` based on:

- expected 2026 year-end close,
- carryover need,
- execution readiness,
- finance priority,
- known scope changes.

### Step 6. Spread the annual ask into months

Create `JobMonthlyPlan2027` by assigning each job's annual ask across the twelve months.

Use the phasing pattern in `Phasing` only as a starting shape. Adjust the monthly spread when a job has:

- a late start,
- a known construction window,
- a site constraint,
- a phased release,
- a procurement lead time.

Rules:

- Do not allocate spend before the draft start month.
- The twelve monthly values for a job must sum to `2027AnnualAsk`.
- Existing and new jobs use the same monthly-plan table.

## Workbook-Ready Formulas

Use Excel structured references. These formulas assume:

- the monthly table is named `JobMonthlyPlan2027`,
- the job table is named `JobMaster2027`,
- the finance table is named `FinanceMonthlyCap`.

The two driver columns in `JobMonthlyPlan2027` are:

- `JobID`
- `MonthNum`

### JobMaster2027 helper formula

Use this in `Remaining2026`:

```excel
=MAX(0, [@[2026Projected]] - [@[2026YTDSpend]])
```

### Lookup columns in JobMonthlyPlan2027

Use these formulas in `JobMonthlyPlan2027`.

`BudgetYear`

```excel
=2027
```

`MonthName`

```excel
=TEXT(DATE([@BudgetYear], [@MonthNum], 1), "mmm")
```

`JobDescription`

```excel
=XLOOKUP([@JobID], JobMaster2027[JobID], JobMaster2027[JobDescription], "")
```

`Region`

```excel
=XLOOKUP([@JobID], JobMaster2027[JobID], JobMaster2027[Region], "")
```

`Site`

```excel
=XLOOKUP([@JobID], JobMaster2027[JobID], JobMaster2027[Site], "")
```

`PM`

```excel
=XLOOKUP([@JobID], JobMaster2027[JobID], JobMaster2027[PM], "")
```

`Category`

```excel
=XLOOKUP([@JobID], JobMaster2027[JobID], JobMaster2027[Category], "")
```

`JobType`

```excel
=XLOOKUP([@JobID], JobMaster2027[JobID], JobMaster2027[JobType], "")
```

`PlanningBucket`

```excel
=XLOOKUP([@JobID], JobMaster2027[JobID], JobMaster2027[PlanningBucket], "")
```

`DraftStartMonth`

```excel
=XLOOKUP([@JobID], JobMaster2027[JobID], JobMaster2027[DraftStartMonth], 1)
```

`PhaseTemplate`

```excel
=XLOOKUP([@JobID], JobMaster2027[JobID], JobMaster2027[PhaseTemplate], "Default")
```

`IncludeInPlan`

```excel
=XLOOKUP([@JobID], JobMaster2027[JobID], JobMaster2027[IncludeInPlan], "N")
```

`AnnualAsk`

```excel
=XLOOKUP([@JobID], JobMaster2027[JobID], JobMaster2027[2027AnnualAsk], 0)
```

`FinanceCapMonth`

```excel
=XLOOKUP(
    1,
    (FinanceMonthlyCap[BudgetYear]=[@BudgetYear])*(FinanceMonthlyCap[MonthNum]=[@MonthNum]),
    FinanceMonthlyCap[FinanceCap],
    0
)
```

### Phase formulas in JobMonthlyPlan2027

For `ManualProjectedSpend`, leave the value blank unless `PhaseTemplate = "Manual"`. Enter a manual month amount only for those jobs.

Use this in `RawWeight`:

```excel
=LET(
    m, [@MonthNum],
    s, [@DraftStartMonth],
    tpl, [@PhaseTemplate],
    defW, {0.04,0.05,0.06,0.12,0.12,0.11,0.10,0.10,0.10,0.07,0.03,0.10},
    frontW, {12,11,10,9,8,7,6,5,4,3,2,1},
    backW, {1,2,3,4,5,6,7,8,9,10,11,12},
    IF(
        OR([@IncludeInPlan]<>"Y", m<s),
        0,
        SWITCH(
            tpl,
            "Default", INDEX(defW, m),
            "FrontLoaded", INDEX(frontW, m),
            "BackLoaded", INDEX(backW, m),
            "Manual", 0,
            INDEX(defW, m)
        )
    )
)
```

Use this in `NormalizedWeight`:

```excel
=LET(
    denom,
        SUMIFS(
            JobMonthlyPlan2027[RawWeight],
            JobMonthlyPlan2027[JobID], [@JobID],
            JobMonthlyPlan2027[BudgetYear], [@BudgetYear]
        ),
    IF(denom=0, 0, [@RawWeight]/denom)
)
```

Use this in `ProjectedSpend`:

```excel
=LET(
    annual, [@AnnualAsk],
    job, [@JobID],
    yr, [@BudgetYear],
    m, [@MonthNum],
    tpl, [@PhaseTemplate],
    rawW, [@RawWeight],
    normW, [@NormalizedWeight],
    lastActiveMonth,
        MAXIFS(
            JobMonthlyPlan2027[MonthNum],
            JobMonthlyPlan2027[JobID], job,
            JobMonthlyPlan2027[BudgetYear], yr,
            JobMonthlyPlan2027[RawWeight], ">0"
        ),
    priorRounded,
        SUMPRODUCT(
            (JobMonthlyPlan2027[JobID]=job)*
            (JobMonthlyPlan2027[BudgetYear]=yr)*
            (JobMonthlyPlan2027[MonthNum]<m)*
            ROUND(JobMonthlyPlan2027[AnnualAsk]*JobMonthlyPlan2027[NormalizedWeight],0)
        ),
    IF(
        tpl="Manual",
        N([@ManualProjectedSpend]),
        IF(
            rawW=0,
            0,
            IF(
                m=lastActiveMonth,
                annual-priorRounded,
                ROUND(annual*normW,0)
            )
        )
    )
)
```

This formula keeps the non-manual month rows rounded while forcing the last active month for each job to absorb the rounding remainder, so the job total ties back to `2027AnnualAsk`.

### Recommended phase template meanings

Use these exact template meanings:

- `Default`: reuse the same month-weight shape as the existing `Phasing` baseline.
- `FrontLoaded`: heavier early-month weighting after the draft start month.
- `BackLoaded`: heavier late-month weighting after the draft start month.
- `Manual`: user-entered monthly values in `ManualProjectedSpend`.

### Step 7. Fit to the finance cap

For each month, compare:

- total planned spend from `JobMonthlyPlan2027`
- finance cap from `FinanceMonthlyCap`

If a month is over cap, resolve it by:

1. pushing lower-priority work to later months,
2. reducing candidate carryover,
3. moving immature new jobs out of the month,
4. splitting large jobs into later release waves.

Do not solve cap pressure by hiding detail. Keep the row-level job plan visible.

### Step 8. Publish rollups

Create rollups from `JobMonthlyPlan2027` for:

- month total,
- month by region,
- month by site,
- month by job.

The minimum final output should support:

- `Month -> total planned vs finance cap`
- `Month + Region -> projected spend`
- `Month + Region + Site -> projected spend`
- `Month + Region + Site + Job -> projected spend`

## Control Checks

Use these checks during drafting.

### Monthly fit

For each month:

`FinanceCap - SUM(ProjectedSpend for that month)`

Expected result:

- positive = unused monthly headroom,
- zero = exact fit,
- negative = over cap and must be resolved.

### Job tie-out

For each job:

`SUM(monthly ProjectedSpend) = 2027AnnualAsk`

### Coverage

Every planned spend row should have:

- region,
- site,
- job id,
- job description,
- month,
- job type.

## Output Shape

The final planning table should be row-level and month-grained:

- `MonthName`
- `Region`
- `Site`
- `JobDescription`
- `JobType`
- `ProjectedSpend`

This single output can feed pivots, summaries, and finance review without losing the job-level trail.

## Recommended Operating Sequence

1. Load finance monthly cap.
2. Seed existing jobs from `Analysis`.
3. Add new jobs.
4. Screen with `Ready` and `defer`.
5. Set annual ask by job.
6. Phase into monthly plan.
7. Fit the months to cap.
8. Publish month, region, site, and job rollups.
