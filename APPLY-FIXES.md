# How to produce the updated Pricing.xlsm

The fix can only be made **inside Excel** (it recompiles the macros). The repo's
`Pricing.xlsm` is the **original, unchanged** workbook. Use the one-run patcher to apply
the verified fixes yourself in a couple of minutes.

## 1. One-click: run the patcher (applies the 15 verified line-level fixes)

`vba/fixes/Apply_Fixes.bas` edits the workbook's own VBA modules in place. It only changes
a line when it matches **exactly**, and reports anything it can't match (nothing is forced).

1. **Work on a copy** of `Pricing.xlsm`.
2. Enable: **File ▸ Options ▸ Trust Center ▸ Trust Center Settings ▸ Macro Settings ▸
   ☑ "Trust access to the VBA project object model"**.
3. **Alt+F11** → double-click **VBAProject** and enter your project password to **unlock** it
   for this session (the object model can't edit a locked project).
4. **File ▸ Import File…** → `Apply_Fixes.bas`.
5. Run **`ApplyAllFixes`** (F5). Read the report (it lists `OK` / `SKIP` per fix).
6. **Debug ▸ Compile VBAProject** → if clean, **save** (keep as `.xlsm`).
7. Optionally delete the `zz_Apply_Fixes` module afterwards.

This covers: the 3 round-1 fixes (`FIXES.md`) + the 12 Tier 1 fixes (`FIXES-DEEP.md`).
Any `SKIP` (e.g. a line whose text differs slightly from the export) → apply that one by
hand from `FIXES-DEEP.md`.

## 2. Still manual (by design — needs your judgement)

- **Tier 2** — confirm intent first (`F_Epicor` PASC WM double-import, `J2` `Offset_B`
  default, `J2` heater-band `Exit Do`). See `FIXES-DEEP.md`.
- **Tier 3** — reliability patterns (error-handler/cleanup, `Integer→Long/Double`, release
  the Selenium driver). These are structural; copy the worked example in
  `vba/fixes/HARDENING-EXAMPLE.md` into each handler.
- **External link** — re-point (don't break) `PETRA SPS V-1.0.xlsb` via Data ▸ Edit Links.

## Why I can't hand you a finished file
This review/automation runs in a Linux cloud container with **no Excel**, so I can't open,
unlock and recompile the workbook here; and editing the compiled `vbaProject.bin` directly
would risk corrupting it. The patcher above is the safe equivalent — the edit happens inside
your Excel, which validates and recompiles it.
