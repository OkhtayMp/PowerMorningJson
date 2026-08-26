# PowerTTman — Morning XLSM Extractor

PowerTTman is a lightweight desktop application for extracting
`Site ID` and `Priority` information from `Morning.xlsm`
and generating a clean JSON file.

## Features

- Drag & Drop interface
- Accepts only `Morning.xlsm`
- Reads the `HO To MS Sites` worksheet
- Extracts only:
  - `Site ID`
  - `Priority`
- Live progress display
- Non-blocking background processing
- Clean completion screen
- Cross-platform application
- Windows executable
- Linux executable
- macOS application
- Automatic JSON storage using `platformdirs`

---

## Input

The application expects:

```text
Morning.xlsm