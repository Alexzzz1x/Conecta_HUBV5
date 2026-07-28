# Conecta Hub 2.0

Single-file Python/Tkinter desktop app — SAP GUI automation hub.

## Entrypoint

- `conecta_hub2.py` — main app, creates the Tk root and notebook with 6 tabs.

## Modules

All under the repo root, no package structure. Each exports a `construir_aba_*` factory.

| File | Tab | SAP needed |
|---|---|---|
| `modulo_etiquetas.py` | Gerador de Etiquetas | No |
| `modulo_extrator.py` | Extrator de Seriais / IQ09 | Yes (IQ09) |
| `modulo_Sap.py` | Login SAP | Yes |
| `modulo_bandeirada.py` | Bandeirada | Yes |
| `modulo_extrair_OSME.py` | Extrator OSME | Yes |
| `modulo_conferencia.py` | Conferência | No |

## Build

```powershell
py -m PyInstaller "Conecta Hub 2.0.spec" --clean --noconfirm
```

Output: `dist\Conecta Hub 2.0\Conecta Hub 2.0.exe`

### Spec file gotchas

- `excludes` list has caused bugs — **never add `numpy`, `PIL`, or `pandas`** to it.
- Hidden imports required: `numpy`, `PIL`, `pandas`, `openpyxl`, `reportlab.*`, `qrcode`, `win32com`, `pythoncom`, `tkinter.*`.
- Data files to bundle: all `.vbs`, `.ico`, `.png`, `.xlsx` in the repo root.

### Resource resolution

- Bundled resources (icons, VBS scripts, logo) → `sys._MEIPASS` via `obter_caminho_recurso()`.
- User-editable files (Excel database, config JSONs) → exe directory (`os.path.dirname(sys.executable)`), with fallback to `_MEIPASS`.

## Dependencies

`pandas`, `openpyxl`, `reportlab`, `qrcode`, `pywin32` (win32com), `Pillow`, `numpy`.

## SAP dependency

OSME, Bandeirada, and IQ09 tabs check `verificar_sap_aberto()` (COM `GetObject("SAPGUI")`) before running. If SAP GUI is closed, a messagebox error is shown and the operation is blocked.

## Python

CPython 3.14, Windows-only (COM + cscript + SAP GUI).
