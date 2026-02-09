# Implementation Plan - Wizped Project Organization

This plan outlines the steps to organize the `wizped` project, initialize the Python environment with `uv`, and document the existing Excel structure.

## Goal

To provide a robust, version-controlled repository structure for the `wizped` Excel automation tool, including dependency management and documentation.

## Proposed Changes

### 1. Project Initialization & Structure

- [x] Initialize `uv` for dependency management.
- [ ] Create a `src/` directory for Python scripts.
- [ ] Move `wizped_import.py` to `src/`.
- [ ] Create `macros/` directory for VBA files (exported for version control).
- [ ] Create `assets/` directory for images/resources.
- [ ] Create `.gitignore` optimized for Python and Excel.

### Dependencies

- `pandas`: Data manipulation.
- `openpyxl`: Excel file reading/writing (robust, no need for Excel instance).
- `xlwings`: Interaction with active Excel instance (macros).
- `pdfplumber`: PDF parsing (existing requirement).

### 2. Documentation

- [ ] Create `README.md` with project overview and usage instructions.
- [ ] Create `docs/schema.puml` with the PlantUML diagram of the Excel database.
- [ ] Create `docs/analysis.md` with a deeper analysis of the project.

### 3. Verification

- Verify `uv sync` works.
- Verify Python scripts run from the new structure.
- Verify Git tracks the correct files.

## User Review Required

- None. This is a structural organization task requested by the user.

## Verification Plan

### Automated Tests

- Run `uv run src/wizped_import.py --help` to verify environment.

### Manual Verification

- Check if `pyproject.toml` exists and contains dependencies.
- Check if `schema.png` (rendered from puml) or the puml file exists.
- Inspect the file structure.
