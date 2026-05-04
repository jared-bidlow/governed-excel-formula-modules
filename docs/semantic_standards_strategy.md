# Semantic Crosswalk Reference

This is an archived/reference note for the optional semantic crosswalk files that are still tracked in this repo. It is not part of the current operator workflow and it is not a recommended next step.

The current workflow remains workbook-centered:

```text
Governance Starter workbook
  -> tblBudgetInput
  -> formula modules
  -> Integration Bridge / CSV handoff
  -> review outputs
```

## What Is Tracked

The reference crosswalk keeps small public-safe starter tables and formulas:

- curated namespace rows for `rec`, `brick`, and workbook-local `gef` labels,
- a small class map for common REC and Brick labels,
- a small relationship map for location, composition, feeds, points, and project impact,
- project and asset mapping starter tables,
- an issue screen,
- a flat triple-shaped review queue.

These files exist for audit coverage and private extension. They do not change the default planning workbook, `tblBudgetInput`, the Integration Bridge, or asset finance calculations.

## What Is Not Included

The public template does not import full REC or Brick ontology dumps. It does not expose hundreds of ontology classes. It does not implement Azure Digital Twins, Fabric graph, RDF, Turtle, or JSON-LD export as a complete integration.

`Ontology.JSONLD_EXPORT_HELP` is guidance only. `Ontology.TRIPLE_EXPORT_QUEUE` is a reviewable workbook table, not a deployed graph pipeline.

## Operator Boundary

Use normal planning unless semantic mapping is explicitly in scope for a private workbook copy. Do not treat this reference note as a current operating path.

## Public Safety

Do not commit tenant URLs, private endpoints, credentials, database names, local workbook paths, real building identifiers, or full ontology source files. The only namespace URLs committed for this slice are official public REC and Brick namespace identifiers in the starter namespace table.
