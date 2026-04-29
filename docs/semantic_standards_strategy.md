# Semantic Standards Strategy

The semantic crosswalk is optional. It is for teams that want capital-planning rows to carry digital-twin-ready context without turning the workbook into a graph database or importing full ontology files.

## Role split

Use RealEstateCore (REC) for buildings, rooms, spaces, real-estate context, and generic facility assets.

Use Brick for equipment, HVAC, electrical, lighting, fire-safety systems, points, sensors, meters, setpoints, and commands.

The workbook remains the transparent planning and review surface. The SemanticTwin edition adds a small crosswalk from workbook project and asset rows to semantic classes and relationships.

## What is included

The first semantic slice includes:

- a curated namespace table for `rec`, `brick`, and workbook-local `gef` predicates,
- a curated class map for common REC and Brick concepts,
- a curated relationship map for location, composition, feeds, points, and project impact,
- project and asset semantic mapping tables,
- an ontology issue screen,
- a simple triple export queue.

## What is not included

The public template does not import full REC or Brick ontology dumps. It does not expose hundreds of ontology classes. It does not implement Azure Digital Twins, Fabric graph, RDF, Turtle, or JSON-LD export as a complete integration.

`Ontology.JSONLD_EXPORT_HELP` is guidance only. `Ontology.TRIPLE_EXPORT_QUEUE` is the reviewable workbook contract for future graph, Fabric, or digital-twin workflows.

## Operator flow

Start with normal planning unless semantic mapping is explicitly in scope. Use the `SemanticTwin` edition only after the planning and optional asset workflow are understood.

Recommended flow:

1. Review source health and planning outputs.
2. If assets are needed, use Asset Hub.
3. If digital-twin context is needed, open Semantic Map Hub.
4. Map `ProjectKey` or `AssetId` to a semantic subject, relationship, object, and REC/Brick class.
5. Review ontology issues.
6. Use the triple export queue as the public-safe handoff table.

## Public safety

Do not commit tenant URLs, private endpoints, credentials, database names, local workbook paths, real building identifiers, or full ontology source files. The only namespace URLs committed for this slice are official public REC and Brick namespace identifiers in the starter namespace table.
