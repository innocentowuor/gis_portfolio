# GIS & Spatial Data Portfolio

A public-safe portfolio of GIS automation, spatial data pipelines, web mapping, and public/environmental health analysis projects.

## About

I am a Geospatial Analyst based in British Columbia with training in Biology, Global Public Health, and Geographic Information Systems. My work focuses on using spatial data, automation, and mapping to support public and environmental health decision-making.

This portfolio highlights examples of my work in spatial ETL, ArcGIS automation, geodatabase management, web GIS workflows, reporting tools, and data storytelling.

## Portfolio Focus

- GIS automation and Python scripting
- ArcPy and ArcGIS Pro workflows
- ArcGIS Online, Survey123, Dashboards, and Experience Builder
- Spatial ETL and schema standardization
- Geodatabase management, QA/QC, and data cleaning
- Web mapping and spatial reporting
- Public health and environmental data storytelling

## Featured Projects

| Project | Focus | Tools |
|---|---|---|
| [AOI Spatial Screening Report](projects/ 01_aoi_screening_tool/aoi_screening_tool.py) | Automated spatial overlay, summary reporting, and map export workflow | Python, ArcPy, ArcGIS Online, Excel |
| [AGOL Feature Layer Backup Tool](projects/ 02_weekly_arconline_backups/) | Downloading and backing up ArcGIS Online feature layers into a local geodatabase | Python, ArcPy, ArcGIS API for Python |
| [Spatial Observation ETL Pipeline](projects/ 03_stewardship_baseline_observation_tool/) | Standardizing multiple field observation datasets into one analysis-ready schema | Python, ArcPy, pandas, ArcGIS Online |

## Repository Structure

```text
gis-spatial-data-portfolio/
│
├── projects/
│   ├── 01_aoi_screening_tool/
│   ├── 02_weekly_arconline_backups/
│   ├── 03_stewardship_baseline_observation_tool/
│
├── reusable_gis_functions/
│   ├── agol_tools.py
│   ├── arcpy_gdb_tools.py
│   ├── spatial_overlay_tools.py
│   ├── schema_mapping_tools.py
│   ├── excel_report_tools.py
│   └── validation_tools.py
│
├── docs/
│   ├── privacy_and_sanitization.md
│   ├── technical_stack.md
│   └── project_index.md
│
├── assets/
│   ├── images/
│   ├── diagrams/
│   └── screenshots/
│
├── README.md
├── requirements.txt
├── environment.yml
├── .gitignore
└── LICENSE
