let
    Source = Excel.CurrentWorkbook(){[Name = "tblPlanningTable"]}[Content]
in
    Source
