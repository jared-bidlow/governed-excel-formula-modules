# Notes And Decision Flow Map

This public-template file shows how to document a controlled workbook decision or notes flow.

It does not describe a real meeting process.

## Generic Flow

```text
Main report row -> note/decision context -> staging table -> validation gates -> controlled writeback
```

## Field Classes

| Class | Meaning | Default route |
|---|---|---|
| Raw operator input | A user types or selects a value. | Validate and stage before writeback. |
| Resolved helper | Formula output derived from raw input and source records. | Do not manually overwrite. |
| Apply gate | Boolean/status field that controls whether a staged change can apply. | Fail closed when ambiguous. |
| Direct writeback target | Workbook source field updated by a script or controlled process. | Require scenario coverage and rollback notes. |

## Public Template Boundary

Use fake field names when publishing examples. Keep real notes, meeting text, project records, and employee names out of the repo.
