# MidiBill3 (VB6)

MIDI billentyűzet-router: egy bejövő MIDI jelet **4 kimeneti csatornára**
szoroz szét egyszerre, mindegyiken önálló hangszínnel, hangerővel, pannal
és vibratóval — vagyis egyetlen billentyűlenyomásból akár 4 hangszer
szólal meg egyidejűleg.

## Állapot

Ez a projekt eredeti, **Visual Basic 6**-ban íródott verziója. Jelenleg is
aktívan használt, éles változat. Emellett folyamatban van egy **VB.NET-es
port** is (`MidiBill3` repó), ami hosszabb távon leváltja ezt a verziót —
addig viszont ez a repó tartja karban a jelenleg ténylegesen futó
programot.

## Funkciók

- MIDI bemenetről érkező jelek szétosztása 4 kimeneti csatornára,
  csatornánként önálló hangszínnel, hangerővel, pannal és vibratóval
- Csatornánkénti be-/kikapcsolás (aktív hangszerek), automatikus
  "csatlakozás" a már lenyomva tartott hangokhoz visszakapcsoláskor
- Preset mentés/betöltés (`.mb3` fájlok), automatikus induláskori
  betöltés / kilépéskori mentés
- "Csend legyen!" vészleállító gomb
- Hibatűrő MIDI-eszközkezelés (eszköz menet közbeni eltávolítása nem
  omlasztja össze a programot)

## Fejlesztői környezet

- Visual Basic 6 (SP6)
- Szükséges ActiveX-vezérlők (regisztrálva legyenek a rendszeren):
  - `MSComCtlLib` (Microsoft Common Controls – csúszkák)
  - `MSComDlg.CommonDialog`
  - `midiio2k.ocx`, `midifl2k.ocx` (Mabry Software – MIDI I/O)

## Projekt felépítés

```
MidiBill.frm      – a teljes UI és üzleti logika
MIDCONST.bas      – MIDI-konstansok (üzenettípusok, kontrollerszámok)
MidiBill.frx      – bináris erőforrások (a .frm-hez tartoznak)
MidiBill.vbp      – VB6 projektfájl
Adatok/           – elmentett hangszer-preset könyvtár (.mb3 fájlok)
Kezdet.mb3        – aktuális, élő beállítás (git által figyelmen kívül hagyva)
```

## Kapcsolódó projekt

A `MidiBill3` (VB.NET) repó tartalmazza a folyamatban lévő portolást,
amely ugyanezt a funkcionalitást hivatott nyújtani modern .NET/WinForms
alapokon, NAudio.Midi-vel a MIDI-rétegen.
