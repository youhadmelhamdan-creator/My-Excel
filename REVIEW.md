# Pricing.xlsm — Review & Findings

A deep review of the macro‑enabled workbook `Pricing.xlsm` (Petra Engineering HVAC
pricing / selection engine) and its VBA project.

## 1. File integrity

| Property | Value |
|---|---|
| Type | Valid OpenXML macro‑enabled workbook (`.xlsm`) — not corrupt |
| Size | ~3.8 MB |
| Worksheets | 51 |
| Defined names | 256 |
| VBA modules | 63 (~16,900 lines, ~460 procedures) |
| Author | Yousef Hamdan |
| Malware scan (olevba) | No malicious code. The flagged keywords (`Shell`, `Kill`, `CreateObject`, `Environ`, IP lookup, …) are all legitimate features of this tool. |

Composition driving the file size: **523 EMF vector images + 7 embedded Word
documents + 1 embedded Office Store add‑in**.

## 2. What the macros do

- **`Workbook_Open` (ThisWorkbook):** unprotect → restyle each product sheet's
  buttons/boxes (`Start_Up`) → mark ~24 helper sheets `xlSheetVeryHidden` →
  re‑protect with password `"3141"` → save.
- **Licensing / auto‑update (`Z1_Protection_IP_Update.CheckUser_IP`):** opens
  `Z:\…\Version.xlsm`, reads version + subscription plan + IP, forces an update
  (rebuilds a fresh workbook) when the version differs, and blocks use unless the
  plan is one of ADMIN/GOLD/VIP/SILVER/BRONZE.
- **Web automation (`A1_ElectricalDataCalculator`):** SeleniumBasic + `msedgedriver.exe`
  drive the internal eDesign portal (`edesign.petra-eng.com.jo:5000`) to compute
  electrical data, logging in with the user's licensed credentials.
- **Pricing / output engines:** PPH/PAH/Chiller/FCU quick pricing, offer creation &
  styling, Epicor export, documentation/drawing generation, Schneider submittal.

## 3. Findings

### Confirmed bugs (fixed copies in `vba/fixes/`)

| # | Module | Location | Problem |
|---|---|---|---|
| 1 | `D1_Style_Sheets.bas` | `Sheet_Design_ED` | `CF_Brand.Top` set twice (`520` then `560`) — first assignment is dead code; one button may be mispositioned. |
| 2 | `Z1_Protection_IP_Update.bas` | `CheckUser_IP`, ~line 479 | `Workbooks("Version").Close` — the open book is `Version.xlsm`; the name without extension raises run‑time error 9. Because `On Error GoTo Internet` is still active, the unauthorized path shows a misleading **"problem in network"** message instead of the authorization warning. |
| 3 | `A1_ElectricalDataCalculator.bas` | line 3 | `ch` declared `As Selenium.ChromeDriver` but assigned `New Selenium.EdgeDriver` / `ch.Start "edge"` — type inconsistency that can raise a type‑mismatch. |

### Code‑quality / robustness (recommendations — not auto‑applied)

- **`Option Explicit` missing in 62 of 63 modules.** Biggest systemic risk: a typo
  silently becomes an empty `Variant`. Example: `Workbook_Open` assigns `u, v, w, z, Y`
  that were never declared *and* never used. The `Dim a, b, c As String` pattern
  (only the last variable is typed) is everywhere. Adding `Option Explicit` is a
  refactor (every undeclared variable must then be declared) — do it module by module.
- **`On Error Resume Next`** at the top of several click handlers / readers hides real
  failures and can leave Excel stuck in `Calculation = xlManual` / `ScreenUpdating = False`.
- **Heavy use of bare `End`** to terminate — resets all state; can leave `Version.xlsm`
  open or calc in manual mode on an error path.
- **120 `.Select` / `.Activate` calls** — slower and more fragile than direct references.

### Portability (will affect other machines)

- **External link is a LIVE dependency — do not just delete it.** The link points to a
  personal path `C:\Users\y-hamdan\Desktop\PETRA SPS V-1.0.xlsb`, but the defined names
  `NO_DATA` (`[1]!NO_EDATA[DATA]`) and `Project_List_data` (`[1]!Project_List_Data_Viwer[NO.]`),
  used by the PAH pricing/data sheets, reference tables inside it. **Re‑point** the link to a
  stable shared location (Data ▸ Edit Links ▸ Change Source); breaking it would turn those
  names into `#REF!`.
- Hardcoded mapped‑drive / UNC paths throughout: `Z:\ksa\Yousef Hamdan\…`,
  `\\srv-fileserver\PetraCatalogs`, `C:\Users\Yousef Hamdan\Desktop\Yousef.pdf`. Everything
  breaks if `Z:` isn't mapped or the username differs. Consider a single config sheet/cell
  for base paths.

### Security notes (design, not vulnerabilities)

- Protection password **`"3141"` is hardcoded in plaintext in the VBA (7×)** and is
  trivially recoverable — treat it as light gating, not real protection. The whole
  licensing / IP gate is client‑side VBA and can be bypassed by disabling macros.
- **Good:** eDesign credentials are *not* hardcoded — they're read from the licensing
  record. **But** they land in plaintext in the hidden `Naming` sheet (e.g. cell `A44`);
  anyone who unhides it (password `3141`) can read a user's portal password. Consider
  clearing that cell after the Selenium login completes.

## 4. Why the workbook binary was not auto‑edited

The VBA is stored as compiled p‑code inside `xl/vbaProject.bin`. Editing that source
outside Excel either (a) has no runtime effect — Excel runs the cached p‑code — or
(b) risks corrupting a complex file (523 images, ActiveX controls, OLE embeddings).
Re‑saving through LibreOffice would likely mangle the ActiveX/OLE content. Therefore the
fixes are delivered as **ready‑to‑import corrected modules** (`vba/fixes/`) plus the
step‑by‑step guide in **`FIXES.md`**, to be applied in Excel's VBA editor. The full VBA
source is also exported to `vba/source/` so it can be version‑controlled and diffed.
