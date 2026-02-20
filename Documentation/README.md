# RMP Web App

> This application is in active development and will change frequently.

## Architecture

Single-page ArcGIS JS API 4.30 application using AMD/Dojo modules. No build step.

### Module Graph

| Module | Role |
|---|---|
| `app.js` | Main orchestrator – wizard flow, AOI selection, analysis pipeline |
| `config-helpers.js` | Config loader, utility functions |
| `map-utils.js` | Screenshot capture, symbology helpers |
| `query-engine.js` | Query execution, paging, coverage stats, per-feature tables |
| `summary-engine.js` | **New** – Attribute summary generation with plugin overrides |
| `final-report.js` | Report HTML generation (3 builders: progressive, final, background) |
| `visual-report.js` | Per-layer screenshot generation for visual report |
| `feature-picker.js` | Interactive feature selection |
| `search.js` | Address/coordinate search |
| `upload-aoi.js` | Shapefile/GeoJSON upload for AOI |

### Config-Driven Layer Registry

All layers are defined in `config.json`. Key sections:

- **`reportLayers[]`** – Data layers analyzed in reports. Each entry has:
  - `title`, `url`, `tier` (1/2/3), `category`, `serviceType` (feature/map/image)
  - Optional: `summaryPlugin` (named plugin for attribute summaries), `useServiceRenderer`, `alwaysVisible`, `imageService`, `renderingRule`
- **`selectionLayers[]`** – Interactive selection layers (PLSS, permits). Each entry has `title`, `url`, `visible`, `group`.
- **`referenceLayers`** – Infrastructure URLs (PLSS state boundary, SMA, states, counties, geocode).
- **`symbology.presets`** – Renderer definitions for selection, report, AOI display.
- **`map`** – Basemap, center, zoom defaults.
- **`report`** – Query batch sizes, page sizes, AOI thresholds.

### Adding a New Layer

1. Add an entry to `reportLayers` in `config.json` with `title`, `url`, `tier`, `category`, `serviceType`.
2. The layer is automatically: queried during analysis, symbolized in the map (with thickened borders), categorized into the correct report section, and summarized using generic auto-classification or a named plugin.
3. Optionally set `summaryPlugin` to a registered plugin name (see `summary-engine.js` for the list).
4. Alternatively, use the **Admin UI** at `admin.html` to manage layers via a web interface.

### Summary Engine Plugins

The `summary-engine.js` module replaces the original ~870-line if/else chain with a plugin registry. Built-in plugins:

`vri`, `critical-habitat`, `grazing-allotments`, `wilderness`, `acec`, `wild-horse-burro`, `ungulate-migration`, `mlrs-row`, `land-use-plan`, `nlcs`, `locatable-minerals`, `timber`, `usfws-regions`, `grazing-pastures`, `oil-gas`, `recreation-sites`, `lwcf`, `eplanning`, `fire-perimeters`, `admin-units`

Plugins are matched by: (1) explicit `summaryPlugin` field in config, (2) title-based auto-match, (3) generic auto-classifier fallback.

### Admin UI

`admin.html` is a self-contained admin page (no dependencies). It connects to the GitHub API using a PAT to:

- View/search/filter all layers by category and tier
- Add, edit, or delete report and selection layers
- Edit reference layer URLs
- Run service health checks (pings each endpoint)
- Commits changes directly to `config.json` in the repo

### Performance Optimizations

- **Export query carry-forward**: Real query objects built during screening are reused during report generation (eliminates duplicate queries)
- **Unified FeatureLayer cache**: `processOneTarget()` uses `queryEngine.getCachedLayer()` instead of creating orphan FeatureLayer instances
- **Intersection caching**: Per-feature intersection results are cached between `computeLayerCoverageStats` and `buildPerFeatureTable`
- **O(1) OID deduplication**: Chunked queries use a `Set` instead of `O(n²)` linear scan
- **Deferred coverage stats**: Coverage stats fire as non-blocking promises during screenshot capture, collected afterward
- **Reduced sleep timers**: Hard sleeps cut from 1500-3500ms to 500-800ms with reactive 150ms settle buffers
- **Category-based bucketing**: O(1) config lookup replaces regex title matching for report categorization
