# Kiwi-Änderungen am sdkjs-Fork

Dieses Verzeichnis ist der Fork von [ONLYOFFICE/sdkjs](https://github.com/ONLYOFFICE/sdkjs)
auf Commit `72b0421c0b` (Merge `release/v9.4.0` in `master`). Die Datei
sagt, wo der Delta liegt, damit ein Versions-Bump ihn findet.

Die Konvention entspricht `onlyoffice-documentserver/KIWI-CHANGES.md`
(AGPL §13 Corresponding Source). Solange das gepatchte `sdk-all.js`
per `docker cp` in einen laufenden Container gelegt wird, beschreibt
jene Datei die laufende Version nicht. Vor dem ersten Deploy muss der
sdkjs-Build ins Overlay-Dockerfile, und der Absatz dort nachgezogen
werden.

## Build und Deploy

`build-sdk-all.py` ist der ganze Build: eine Konkatenation der in
`configs/slide.json` gelisteten Dateien, in Reihenfolge, ohne Toolchain
und pro Commit byte-identisch.

```
python3 build-sdk-all.py -o /tmp/sdk-all.js
```

Der **Deploy-Weg** ist das Overlay-Dockerfile
(`onlyoffice-documentserver`), das diesen Fork als Submodul führt und
das Skript zur Build-Zeit ruft. `docker cp` ist ein Dev-Trick zum
Ausprobieren, kein Deploy: ein per `docker cp` bespielter Container
entspricht keinem Repo, und damit beschreibt die veröffentlichte
Quelle die laufende Version nicht (AGPL §13).

## sdk-all.bin

Editor-Save läuft durch DoctRenderer (`apply_changes`). Der Prozess
lädt `slide/sdk-all.bin`, wenn die Datei da ist — ein Snapshot des
ungepatchten SDK. Nach einem Bundle-Deploy `sdk-all.bin` und
`sdk-all.cache` löschen, sonst laufen weder Write noch Leiste.

Nicht `sdk-all-min.js` ersetzen (offizielles 2015er Min). Concat von
`configs/slide.json` nach `/tmp`, dann nach
`/var/www/onlyoffice/documentserver/sdkjs/slide/sdk-all.js`.

## Modell: ein Abschnitt ist ein Schnitt, keine Menge

Ein Abschnitt ist ein Schnitt in der Folienreihenfolge — genau das,
was die Datei hergibt. Upstream hat dafür bereits `CPrSection` mit
`startIndex`; der Fork legt kein `slides`-Array daneben.

```
sectionOf(i) = letzter Schnitt mit start <= i
```

Damit gibt es keine Mitgliedschaftsmengen, keine Adoption, keine
Waisen, keine nicht-zusammenhängenden Abschnitte. Zustände, die die
Datei nicht darstellen kann, sind nicht konstruierbar.

### `normalise()` — das einzige Fixup

Eine totale Funktion über den Abschnittsvektor, aufgerufen nach jeder
Folienbewegung:

1. nach `startIndex` sortieren (stabil, Array-Position als
   Tiebreaker — `Array.prototype.sort` ist nicht überall stabil),
2. in `[0, count]` klemmen,
3. den ersten Schnitt auf 0 pinnen.

**Es wird nichts verworfen.** Ein Abschnitt, dessen Folien alle
weggezogen wurden, behält seinen Schnitt und besitzt eben nichts —
genau das speichert PowerPoint auch (`<p14:section>` mit leerer
`sldIdLst`). Nur so kann ein versehentlich geleerter Abschnitt wieder
befüllt werden, statt neu angelegt werden zu müssen.

Liegen mehrere Schnitte auf demselben Index, besitzt der **letzte** die
folgenden Folien (`sectionOf` nimmt den letzten Schnitt ≤ i).

### Zwei Arten von Schreibzugriff, nie gemischt

Das ist die Regel, an der „Bei der Bearbeitung ist ein Fehler
aufgetreten“ hing.

| Art | Wann | Wie |
| --- | --- | --- |
| **abgeleitet** | Eine Folie wurde verschoben / eingefügt / gelöscht | `section.startIndex = v` direkt, **ohne History**. Die Folienänderung steht bereits in der History; `syncFromChange` leitet bei Undo/Redo neu ab. |
| **absichtlich** | Abschnitt anlegen, umbenennen, verschieben | `StartAction` / `FinalizeAction` um die Upstream-Setter `setName` / `setStartIndex` / `setGuid` — undoable. |

`CPrSection.prototype.setName/setStartIndex/setGuid` rufen
`History.Add(...)` (`Presentation.js:235–246`). Diese Setter außerhalb
einer offenen Action zu rufen, löst den Fehlerdialog aus. `setStart()`
im Modul umgeht sie deshalb bewusst und sagt das im Kommentar.

### Der Drop-Resolver: die einzige Rückrichtung View → Modell

Das Modell weiß nicht, was eine Folienleiste ist. Es fragt:

```js
PresentationSections.setDropResolver({ownerFor, afterShift})
```

Die Leiste registriert das beim Laden. `ownerFor(presentation, pos)`
liest den Drop vom Manager (`sectionGap`), `afterShift` lässt die Leiste
neu layouten. Damit gibt es kein modulglobales „der letzte Drop war…"
mehr, das veralten könnte, und der Pfad deckt **jeden** Aufrufer von
`shiftSlides` ab — auch `DublicateSlide`, was ein Hook in einem
Maushandler nie konnte.

### Die Naht: ein Drop-Ziel nennt den Besitzer, nicht nur den Index

Einfügeindex `k` heißt zugleich „Ende von A“ und „Anfang von B“ — und
mit leeren Abschnitten können mehrere Schnitte auf `k` liegen. Ein
Boolean reicht dafür nicht. Jedes Drop-Ziel trägt deshalb die
**Section**, die die Folie bekommen soll:

```
gaps[] = {insertAt, owner, along, lo, hi}
```

`shiftForInsert(pos, n, owner)` schiebt jeden Schnitt **hinter** `pos`
und jeden Schnitt **auf** `pos`, der in der Reihenfolge nach `owner`
kommt. Ohne genannten Besitzer gilt: *die neue Folie kommt in den
Abschnitt der Folie, der sie folgt* — das ist, was Duplizieren,
Einfügen und Paste wollen (und war der Grund, warum das Duplikat einer
einzelnen Folie im nächsten Abschnitt landete).

`dropTarget()` ist die einzige Stelle, an der aus einer Position ein
Ziel wird, und das Layout, das die Linie zeichnet, ist dasselbe.

### Upstreams `addSection` ist nicht registriert

`CPresentation.addSection` existiert upstream und schreibt
`historyitem_Presentation_AddSection` — aber weder `changesFactory`
noch `drawingContentChanges` kennen den Typ, das Array kommt bei Undo,
Redo und fremden Änderungen also `undefined` zurück.
`registerSectionChanges()` in `install()` trägt die beiden Einträge
nach; danach funktioniert die Upstream-Methode wie geschrieben, und der
Fork braucht keinen zweiten Weg, eine Section anzulegen.

`inAction()` fängt außerdem. Eine Exception, die aus einem
Maushandler entkommt, landet im `window.onerror` des Editors: Dialog
„Bei der Arbeit mit dem Dokument ist ein Fehler aufgetreten“,
Co-Editing getrennt, Editor auf View-Only. Eine verlorene Section ist
reparabel, eine verlorene Sitzung nicht; der Stack geht mit Präfix
`kiwi sections:` in die Konsole.

## Hooks (rebase bei Versions-Bump)

| Datei | Was |
| --- | --- |
| `configs/slide.json` | `PresentationSections.js`, `PresentationThemeRemount.js`, `PresentationSectionRail.js` nach `Presentation.js` |
| `common/Drawings/Format/Format.js` | `CPres.fromStream` case 7: `PresentationSections.read(s)` |
| `common/Shapes/SerializeWriter.js` | `WritePresentation`: `PresentationSections.write` — kein Remount, kein Adopt |
| `common/Shapes/Serialize.js` | `Load`: `PresentationSections.apply` |
| `slide/Editor/Format/Presentation.js` | **kein Delta.** `insertSlide`, `removeSlideByObject`, `shiftSlides` werden in `PresentationSections.hookPresentation` als Prototype-Wrapper installiert |
| `CChangesDrawingsContentPresentation` | `hookHistoryChanges` wrappt `Load` / `Redo` / `Undo` — steht in keinem git diff; `write()` warnt, wenn der Hook fehlt |
| `slide/Drawing/DrawingDocument.js` | **14** Hooks, s.u. (+153 / −15 gegen HEAD) |

### `apply-kiwi-rail.py` ist der Rebase-Check

`DrawingDocument.js` ist die einzige Upstream-Datei mit nennenswertem
Delta. Der Patch liegt nicht als Diff vor, sondern als 14 verankerte
Hooks in `apply-kiwi-rail.py`. Jeder Anker muss **genau einmal**
matchen, sonst bricht das Skript ab; ein zweiter Lauf auf einer bereits
gepatchten Datei wird ebenfalls verweigert. Nach einem Upstream-Bump:

```
git checkout slide/Drawing/DrawingDocument.js
python3 apply-kiwi-rail.py slide/Drawing/DrawingDocument.js
```

Das Skript nennt den Hook, dessen Anker gewandert ist — statt eines
Merge-Konflikts über 187 Zeilen bekommt man einen Namen.

Die Hooks, in Reihenfolge:

```
 1 onMouseDown                 8 OnUpdateOverlay draw
 2 onMouseMove header track    9 insertion line horizontal extent
 3 onMouseMove hover          10 insertion line vertical
 4 onMouseUp header track     11 ConvertCoords2
 5 onMouseLeave clear hover   12 CalculatePlaces layout
 6 manager helpers            13 CalculatePlaces slots
 7 OnPaint skip hidden        14 CheckSizes extra
```

1–5 sind Handler-Rümpfe: sie werden in `initThumbnails` als Closures pro
Instanz gebunden, da kommt kein Prototype-Wrapper ran. Sechs frühere
Hooks sind weg — `MouseDownTrack.*` (der Zustand lag schon als
`sectionGap` am Manager), die Feld-Deklarationen (`undefined` ist falsy),
`first visible page helper` (die zwei Aufrufer machen es inline) und
`onMouseUp shiftSlides`, der als Drop-Resolver ins Modell gewandert ist
(s.u.).

Jeder Hook ist eine Ein-Zeilen-Delegation in die Host-Adapter-Sektion
von `PresentationSectionRail.js` (`hostLayout`, `hostDraw`,
`hostDropIndex`, `hostMouseDown`, `hostMouseMove`, `hostMouseUp`,
`hostRevealDrop`, `hostClearHover`) — das ist der einzige Teil, der
`CThumbnailsManager` kennt.

Zwei Hooks sind Fehlerbehebungen, keine Erweiterungen:

- **`ResetSimple`** delegiert an `SetPosition`, statt `Position` roh zu
  setzen — sonst verliert ein kurzer Drag `beforeCut` und landet im
  falschen Abschnitt.
- **`first visible page helper`**: `drawThumbnailsInsertionLine` liest
  `m_arrPages[0].left/right`. Eine eingeklappte Folie behält ihre
  x-Spalte und bekommt `top === bottom` — **kein `-10000`-Sentinel**,
  damit jeder Upstream-Leser von `m_arrPages[0]` heil bleibt.

## UI

**Die Leiste zeichnet auf `m_oThumbnailsBack`, und die liegt UNTER der
Thumbnail-Canvas** (`initThumbnails` hängt sie zuerst in den
Container). Alles, was über einer Folie liegt, ist unsichtbar. Die
Balken überleben, weil sie in der Lücke sitzen; jedes andere Element
muss in den Rand.

- **Balken**: ein Header pro Abschnitt — auch pro *leerem*. Pille mit
  Haarlinie, kiwi-Tick (`#F29C55`) an der Vorderkante, Chevron,
  Titel, rechts die Folienanzahl bzw. kursiv „leer“. `colors(skin)` hat
  Nicht-Hex-Fallbacks, damit kein Custom-Skin den Balken unsichtbar
  macht.
- **`+`**: ein runder Button im Rand neben der Folie, auf deren
  Vorderkante, plus ein Strich quer durch die Leiste, der zeigt, wo der
  Schnitt landet. Hover-Zone ist die ganze Zeile. Nur dort, wo
  `canAddSection` es erlaubt.
- **Titel**: `titleOf` schlägt `section.name` in
  `document.options.kiwiUnitTitles` nach. **Nicht** in
  `editorConfig.customization` — das verarbeitet die web-apps-Oberfläche
  und es erreicht das SDK nie (`asc_CDocInfo` hat kein solches Feld).
- **Rechte**: `canAddSection` / `canRenameSection` / `moveSection`
  prüfen `CanEdit()`, und die Affordanzen verschwinden mit — im
  View-Only-Modus ist das `+` nicht nur wirkungslos, es ist weg.
- **Umbenennen**: erlaubt, wenn der Name *keine* Unit-Id ist — eine
  Unit-Section umzubenennen würde den Titel-Lookup und die Identität
  des Decks brechen. Das Prädikat überlebt einen Reload, ein
  „frisch angelegt“-Flag tat das nicht. `rail.startRename` legt ein
  `<input>` in den Container, schluckt alle Maus- und Tastenevents
  (sonst werden die Anschläge zu Editor-Shortcuts) und fokussiert im
  nächsten Tick, damit der Fokus-Restore des Editors nicht gewinnt.
  **Blur bestätigt nicht** — der Editor entzieht den Fokus bei jedem
  Mausereignis in der Leiste, blur ist also kein Signal, sondern
  Rauschen. Beendet wird mit Enter, Escape, einem Klick woanders hin
  oder wenn die Leiste scrollt. Eine neu angelegte Section startet
  direkt im Umbenennen.
- **Abschnitt verschieben**: `beginTrack` / `moveTrack` / `endTrack`
  auf dem Balken, mit Schwellwert, damit ein Klick kein Drag ist.

### Was nicht geht: der Kontextmenü-Eintrag

Das Rechtsklick-Menü der Folienleiste wird von **web-apps** gerendert;
sdkjs schickt nur `CContextMenuData` mit Flags
(`IsSlideSelect`, `IsSlideHidden`, …). Ein eigener Eintrag „Neuen
Abschnitt davor einfügen“ ginge nur mit einem Patch am vorgebauten
web-apps-Bundle im Docker-Image — ein zweiter Patch-Kanal, minifiziert,
ohne Rebase-Check. Deshalb bleibt es beim Button in der Leiste.

`PresentationThemeRemount.js` hängt Stock-Themes um und wird aus
`BinaryPPTYLoader.Load` gerufen, direkt neben `PresentationSections.apply`
— beides sind Reparaturen am gerade gelesenen Dokument. Nicht aus
`WriteDocument2` rufen: ein Serialisierer ändert nicht, was er
serialisiert. (`slides/split_ooxml.py` verlässt sich darauf: „no donor
here; remountOfficeHijacks in the editor handles those".)

## Module

- `slide/Editor/Format/PresentationSections.js` — Modell. PPTY-Record 7
  (`p14:sectionLst`), Bind (`sldId - 256`, erste Section gewinnt,
  Schnitt = Index der ersten gebundenen Folie), Write remintet
  `256 + Index` über den Bereich. Bei `sectionsUnbound` /
  `sectionsPartial` (PowerPoint-Renumbering) schweigt `write()`,
  statt Zugehörigkeit zu zerstören.
- `slide/Editor/Format/PresentationSectionRail.js` — Leiste. `layout()`
  liefert `{headers, slots, gaps, extra}` in einem Durchlauf;
  `dropTarget()` ist die einzige Positions→Index-Abbildung.
- `slide/Editor/Format/PresentationThemeRemount.js` — Theme-Remount.

Tests:

```
node slide/Editor/Format/PresentationSections.test.js      #  70 checks
node slide/Editor/Format/PresentationSectionRail.test.js   # 103 checks
node slide/Editor/Format/PresentationThemeRemount.test.js  #  3 checks
```

`PresentationSections.test.js` enthält einen handgeschriebenen
PPTY-Writer/Reader und fährt einen echten Roundtrip; der Rail-Test
prüft u.a. die Naht („eine Linienposition pro Ergebnis“), den
Ein-Folien-Abschnitt, den geleerten Abschnitt (Balken bleibt, ist
Drop-Ziel, lässt sich wieder befüllen), eingeklappte Abschnitte und
dass der Chip im Rand und nie in einem Balken liegt.
