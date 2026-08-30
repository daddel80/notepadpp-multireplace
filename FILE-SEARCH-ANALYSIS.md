# File-Search-Subsystem: Tiefenanalyse (Stand 2026-08-30, Arbeitsstand nach 6.0.0.36)

Anlass: guy038-Report (Forum Post 106204, GitHub Issue #133). Geprüft wurden alle vier Suchpfade
End-to-End im aktuellen Arbeitsstand des Repos sowie zum Vergleich der N++-master (fd0add6, 2026-08-22).
Jeder Befund ist im Quelltext verifiziert, Fundstellen sind als Funktion plus ungefähre Zeile angegeben.

---

## 1. Ergebnis in Kürze

Der Kern des guy038-Reports besteht aus zwei bestätigten Fehlern (UTF-16-Behandlung in Find in Files,
fehlende Transparenz über übersprungene Dateien). Die Tiefenanalyse hat darüber hinaus einen
Crash-fähigen Fehler in Find in Docs (Single Mode), eine Integritätslücke beim Rückschreiben in
Replace in Files sowie mehrere Konsistenzbrüche zwischen den vier Suchpfaden ergeben. Strukturell
ist die Hauptursache immer dieselbe: Laden, Binärerkennung und Encoding-Entscheidung liegen nicht in
einer gemeinsamen Pipeline, sondern sind über HiddenSciGuard, Encoding und zwei Aufrufstellen in
MultiReplacePanel verteilt und dort unterschiedlich verdrahtet.

---

## 1.1 Verifikationsstatus (zweite Runde, 2026-08-30)

Jeder Befund wurde vor der Umsetzungsplanung erneut gegen den Quelltext geprueft. Die
plattformneutrale Kernlogik wurde zusaetzlich mit woertlich uebernommenen Code-Passagen als
lauffaehiger Test bewiesen (src/tests/file_search_qa.cpp, g++ -std=c++20
-fsanitize=address,undefined, 15/15 Checks PASS). Win32-Verhalten ist per Microsoft-Dokumentation
belegt, wo noetig.

| Befund | Status | Beweisart |
|---|---|---|
| F1 UTF-16-BOM auf Rohpfad | BESTAETIGT | Code + Test F1a-F1e (loadFile laesst BOM durch, isLikelyBinary faengt ab, "Fi" in Rohbytes 0 Treffer, detectEncoding haette UTF16 erkannt) |
| F2 replaceListData[0] OOB in Docs Single Mode | BESTAETIGT | Code: Bail-Guards nur im Listen-Zweig (ca. 9523), collect setzt findTextW aus replaceListData[critIdx] (ca. 9622), Single Mode ruft collect(0, ...) (ca. 9672) |
| F3 UTF-16 ohne BOM als binaer verworfen | BESTAETIGT | Code + Test F3a/F3b; N++-Gegenseite (LE-Erkennung via IsTextUnicode) im N++-master verifiziert |
| F4 stille Skips, Zaehler nie angezeigt | BESTAETIGT | Grep ueber gesamtes src: _skipped*-Zaehler ausserhalb HiddenSciGuard.h nirgends gelesen; odd-length-Skip per Test F4a-F4c bewiesen |
| K1 fehlende gemeinsame Pipeline | BESTAETIGT | folgt aus F1/F3/F4 (strukturell) |
| K2 zwei Filter-Engines | BESTAETIGT (intern); Teilfrage OFFEN | Sonderfall *.* in matchesDocFilter (ca. 1233) vs. kein Sonderfall in matchPath: Code. PathMatchSpecW-Verhalten fuer endungslose Namen: MS-Doku unspezifisch ("MS-DOS wildcard match type"), nur per Windows-Test klaerbar. N++-Paritaet im Files-Pfad gilt unabhaengig davon (gleiche API, gleicher Default) |
| K3 Hidden-Semantik | BESTAETIGT, verschaerft | Label ist "In hidden folders" (languages.ini Z. 46, ctrlMap ca. 801) und verspricht damit exakt die N++-Semantik; Implementierung filtert stattdessen versteckte Dateien (matchPath Schritt 1) und prunt keine Ordner (Iteration ca. 9101/9806). Label und Code widersprechen sich |
| K4 Rohpfad-Zweck unklar | BESTAETIGT | Test K4a/K4b: NUL jenseits 8 KB erreicht den Zweig |
| K5 Zero-Length-Treffer verworfen | BESTAETIGT | Alle vier Sammelpfade filtern h.length > 0 (ca. 9423, 9475, 9621, 9958); N++ zaehlt sie (processRange: processed bleibt true, ++nbProcessed) |
| R1 Best-Fit/FFFD beim Rueckschreiben | BESTAETIGT, mit Nuance | flags=0 in wstringToBytes (ca. 240, Kommentar nennt es "Permissive mode"); MS-Doku: ohne WC_NO_BEST_FIT_CHARS Best-Fit/Default-Char, ungueltige Sequenzen werden ohne WC_ERR_INVALID_CHARS still zu U+FFFD. Nuance: Encoding.cpp besitzt bereits roundtripLossless (ca. 89, mit WC_NO_BEST_FIT_CHARS und usedDefault-Check), genutzt aber NUR intern von autoDetectAnsiCodepage, nicht im Write-Back. Der Baustein fuer den Fix existiert also schon |
| R2 writeFile nicht atomar | BESTAETIGT | Code (ofstream trunc, ca. 344) |
| R3 int-Casts > 2 GB | BESTAETIGT | Code-Inspektion (static_cast<int> auf Laengen); Laufzeittest unpraktikabel, Arithmetik eindeutig |
| R4 Enumeration ohne Cancel/Progress | BESTAETIGT | Beide Schleifen pruefen nur _isShuttingDown (ca. 9101-9110, 9806-9822), kein _isCancelRequested, keine Message-Pumpe |
| R6 BOM schaltet Binaercheck ab | BESTAETIGT | Code + Test F1a |
| P1-P4 Paritaets-Entlastungen | BESTAETIGT | N++-master gegengelesen (matchInList, findInOpenedFiles beide DocTabs, dotMatchesNL-Grep) |

Einordnung fuer das Ziel "gleiche Dateizahl wie guy038": Mit Binary-Skip aus ist die gleiche
DATEIMENGE erreichbar (F1+F3+K3 beseitigen die Mengenabweichungen). Identische TREFFERZAHLEN
innerhalb echter Binaerdateien haengen zusaetzlich davon ab, wie nah unsere Dekodierung im
Rohfall an der N++-Ladelogik liegt (uchardet vs. eigene Erkennung); dort sind kleine Deltas
moeglich und beim Design der Option einzuplanen.

---

## 2. Ist-Architektur: die vier Suchpfade

| Schritt | Find All (akt. Doc) | Find in Docs | Find in Files | Replace in Files |
|---|---|---|---|---|
| Quelle | aktiver Scintilla-Buffer | alle offenen Buffer, per NPPM_ACTIVATEDOC durchgeschaltet | Datei von Platte, hidden Scintilla | Datei von Platte, hidden Scintilla |
| Filter | keiner | matchesDocFilter (nur wenn "All Docs" aus) | parseFilter + matchPath | parseFilter + matchPath |
| Encoding | von N++ dekodiert | von N++ dekodiert | loadFile, dann isLikelyBinary-Weiche, dann detectEncoding | loadFile, dann direkt detectEncoding |
| Binärskip | n/a | n/a | NUL in ersten 8 KB ohne BOM | NUL in ersten 8 KB ohne BOM |
| Suchbytes | convertAndExtendW mit Doc-Codepage | dito, je Doc | dito, je Datei (hidden Buffer) | dito |
| Rückschreiben | n/a | n/a | n/a | convertUtf8ToOriginal + writeFile |

Fundstellen: handleFindAllButton (MultiReplacePanel.cpp ca. 9322), handleFindAllInDocsButton (ca. 9511),
handleFindInFiles (ca. 9758), handleReplaceInFiles (ca. 9036), HiddenSciGuard.h (loadFile ca. 264,
shouldSkipAsBinary ca. 249, matchPath ca. 164, parseFilter ca. 127, writeFile ca. 344),
Encoding.cpp (detectEncoding ca. 176, convertBufferToUtf8 ca. 299, convertUtf8ToOriginal ca. 343).

Positiv verifiziert: Suchbytes werden je Datei bzw. je Dokument neu mit der jeweils gebundenen
Codepage erzeugt (convertAndExtendW 2-Arg ruft getCurrentDocCodePage nach dem Binding). Der
HiddenSciGuard ist pro Operation lokal, Zaehler starten sauber bei 0. Die Enumeration von Find in
Files und Replace in Files ist identisch verdrahtet (gleicher Filter, gleiche matchPath-Logik).
ResultDock kappt Anzeigetext pro Treffer bei 2048 Bytes (kMaxHitTextUtf8, ResultDock.cpp ca. 1282),
lange Binaerzeilen sprengen die Anzeige also nicht.

---

## 3. Fehler (muss gefixt werden)

### F1: UTF-16 mit BOM wird in Find in Files als Rohbytes durchsucht
handleFindInFiles (ca. 9905): isLikelyBinary prueft auf irgendein NUL-Byte im gesamten Puffer und
greift VOR detectEncoding. UTF-16-Dateien mit BOM, die loadFile wegen des BOM bewusst durchlaesst,
landen dadurch im ANSI-Rohpfad (SCI_SETCODEPAGE 0, Rohbytes). Ein Muster wie "Fi" trifft in
"F\0i\0" nie. Replace in Files (ca. 9196) hat diese Weiche nicht und dekodiert korrekt ueber
detectEncoding. Damit findet Find in Files Treffer nicht, die Replace in Files ersetzen wuerde:
der schwerste Konsistenzbruch im Subsystem und der Kern des guy038-Reports.

### F2: Find in Docs, Single Mode: Zugriff auf replaceListData[0] ohne Guard
handleFindAllInDocsButton, collect-Lambda (ca. 9622): h.findTextW = replaceListData[critIdx].findText.
Im Single Mode wird collect(0, ...) aufgerufen, der Treffertext aber trotzdem aus der Liste gelesen.
Bei leerer Liste ist das ein Out-of-Bounds-Zugriff auf std::vector (undefiniertes Verhalten,
Crash-Klasse), bei gefuellter Liste ein falsches findTextW (beeinflusst die Trefferzuordnung, vgl.
ResultDock ca. 5290). Die drei Schwesterstellen machen es richtig: Find All nutzt findW (ca. 9480)
bzw. item.findText (ca. 9428), Find in Files nutzt pattW (ca. 9954). Kleiner, isolierter Fix.

### F3: UTF-16 ohne BOM wird als Binaerdatei verworfen
shouldSkipAsBinary laesst nur BOM-Dateien passieren; UTF-16 ohne BOM enthaelt NULs und wird
uebersprungen. N++ erkennt UTF-16 LE ohne BOM (Utf8_16_Read::determineEncoding: Heuristik
Erstbytes plus IsTextUnicode). UTF-16 BE ohne BOM erkennt auch N++ nicht (im N++-Quelltext bewusst
auskommentiert, Begruendung: Erkennung zu schwach). Ziel: LE ohne BOM erkennen und dekodieren,
BE ohne BOM bewusst nicht (N++-Paritaet), beides dokumentieren.

### F4: Stille Skips ohne jede Rueckmeldung
Es gibt fuenf Skip-Gruende, keiner wird gemeldet: binaer (Zaehler existiert, wird nie angezeigt),
Groessenlimit (dito), Encoding-Konvertierung fehlgeschlagen (continue ohne Zaehler, z. B. UTF-16 mit
ungerader Bytezahl, convertBufferToUtf8 ca. 316), Datei nicht lesbar (continue), Read-only bei
Replace (continue, zudem erst NACH dem vollstaendigen Laden geprueft). Der Dock-Header nennt nur
"X hits in Y file(s)" (languages.ini 300-301), ohne "of M searched". Genau daran ist guy038
gescheitert. Minimum: Skip-Zaehler je Grund fuehren und im Header bzw. Status ausgeben, plus
"of M searched" wie bei N++.

### F5: Replace in Files erbt "Find: Search from cursor position" (BESTAETIGT, durch Live-Test)
handleReplaceInFiles fuellt den versteckten Buffer per SCI_ADDTEXT; das laesst den Caret am
Textende stehen. handleReplaceAllButton/replaceAll berechnen den Start ueber computeAllStartPos
mit dem globalen allFromCursorEnabled: Option an plus Wrap aus (Default) bedeutet Start am
Dateiende, Ergebnis 0 Ersetzungen, still. Live reproduziert (cursortest.txt: 0 statt 3
Ersetzungen). Find in Files startet dagegen hart bei 0, nur Replace in Files kippte.

### F6: Selection-Modus leckt in den Datei-Scope (Beifund zu F5)
Find in Files erzwingt ctx.isSelectionMode = false; Replace in Files lief ueber
handleReplaceAllButton/replaceAll, die das Selection-Radio ungefiltert lesen. Mit aktivem
Selection-Radio wuerden Editor-Selektionsbereiche auf fremde Dateien angewendet. Gleiche
Fehlerklasse: Editor-Zustand leckt in den Datei-Scope.

---

## 4. Konsistenzbrueche (Konzeptentscheidung noetig)

### K1: Keine gemeinsame Lade- und Decode-Pipeline (Root Cause hinter F1/F3/F4)
Heute: loadFile im Guard, detectEncoding an zwei Aufrufstellen, isLikelyBinary nur an einer,
Skip-Gruende verstreut. Soll: eine Funktion mit klarem Kontrakt, genutzt von Find in Files UND
Replace in Files, etwa:

    enum class LoadResult { Ok, SkipBinary, SkipTooLarge, SkipUnreadable, SkipUndecodable };
    LoadResult loadTextFile(path, std::string& u8, Encoding::EncodingInfo& enc);

Reihenfolge im Kontrakt: Header lesen, BOM pruefen, UTF-16-Heuristik (LE ohne BOM), erst danach
NUL-Binaercheck, dann Volllesen und Dekodieren. Die Entscheidung faellt genau einmal, beide Pfade
verhalten sich identisch, jeder Skip hat einen benannten Grund (traegt direkt F4). Die reine
Erkennungs- und Konvertierungslogik gehoert dabei in eine testbare Uebersetzungseinheit ohne
Win32-Abhaengigkeit im Kern, damit sie im Container mit g++ -fsanitize=address,undefined pruefbar
ist (Roundtrip-Tests, siehe Abschnitt 7).

### K2: Zwei Filter-Engines mit unterschiedlicher Semantik
Files-Pfad: parseFilter/matchPath (PathMatchSpecW, kein Sonderfall fuer *.*). Docs-Pfad:
matchesDocFilter (eigener Parser, Sonderfall: *.* und * matchen ALLES, auch endungslose Dateien,
MultiReplacePanel.cpp ca. 1231). Je nach PathMatchSpecW-Semantik fuer endungslose Namen behandeln
beide Pfade dieselbe Filterangabe unterschiedlich. Gegen N++ besteht hier KEIN Delta: N++ nutzt
dieselbe API mit demselben Default (matchInList, Common.cpp ca. 1110, getAndValidatePatterns).
Soll: eine Engine fuer beide Pfade; Verhalten fuer endungslose Dateien einmal auf Windows testen
und festschreiben.

### K3: Hidden-Semantik weicht in beide Richtungen von N++ ab
N++: Checkbox heisst sinngemaess "in hidden folders", prunt versteckte ORDNER bei der Rekursion,
nimmt versteckte DATEIEN immer mit (getMatchedFileNames: Hidden-Check nur im Directory-Zweig).
MultiReplace: matchPath filtert versteckte DATEIEN, steigt aber ungefiltert in versteckte Ordner ab
(recursive_directory_iterator ohne Pruning, handleFindInFiles ca. 9806). Ergebnis: MR schliesst
versteckte Dateien aus, die N++ findet, und findet Dateien in versteckten Ordnern, die N++
ausschliesst. Das erklaert Zaehldifferenzen der Art 833 vs. 835 Dateien aus guy038s erstem Test.
Entscheidung noetig: N++-Paritaet (empfohlen, weil die UI-Beschriftung dieselbe Erwartung weckt)
oder bewusst eigene Semantik, dann dokumentieren. Die heutige Mischform ist keins von beidem.

### K4: Der Rohbyte-Zweig hat nach F1 keinen klaren Zweck mehr
Nach dem F1-Fix bleiben im isLikelyBinary-Zweig nur noch Dateien mit NUL jenseits der ersten 8 KB
oder mit hohem Steuerzeichenanteil. Deren Zeilen erscheinen im Dock als Mojibake (Dock ist UTF-8,
Quelle Rohbytes), Suchbytes werden ACP-kodiert. Im Zuge der geplanten Binary-Option sauber
definieren: Skip an bedeutet, der Zweig entfaellt komplett; Skip aus bedeutet bewusste
Rohbyte-Suche wie bei N++ (dann fuer alle Dateien, nicht nur fuer heuristisch erkannte).

### K5: Zero-Length-Regex-Treffer werden verworfen
Alle vier Sammelpfade filtern mit if (h.length > 0). Ein Muster wie ^ oder \b zaehlt in N++ pro
Zeile bzw. pro Grenze, in MultiReplace 0. Fuer die Dock-Anzeige ist das Verwerfen vertretbar, fuer
die ZAEHLUNG ist es eine stille Abweichung, die beim naechsten Vergleichstest der Community wieder
als "MultiReplace findet weniger" auffallen wird. Empfehlung: Treffer mit Laenge 0 zaehlen und als
Positionstreffer anzeigen (N++ macht genau das).

---

## 5. Robustheit (sollte, klar priorisierbar)

### R1: Replace in Files schreibt die ganze Datei durch decode/encode zurueck
Der Rueckweg (utf8ToWString, wstringToBytes mit flags=0, Encoding.cpp ca. 240) erlaubt Best-Fit:
Zeichen, die die Ziel-Codepage nicht kennt, werden still zu Ersatzzeichen. Kaputte
UTF-16-Surrogate werden zu U+FFFD, ungueltige ANSI-Bytes normalisiert. Da die GANZE Datei neu
kodiert wird, passiert das auch weitab der Ersetzungsstellen. Absicherung mit kleinem Eingriff:
vor dem Schreiben Roundtrip-Check encode(decode(original)) == original auf den unveraenderten
Teilen bzw. einmal fuer die Gesamtdatei beim Laden; schlaegt er fehl, Datei ueberspringen und als
SkipUndecodable melden statt still zu veraendern.

### R2: writeFile ist nicht atomar
HiddenSciGuard::writeFile oeffnet mit trunc und schreibt direkt (ca. 344). Absturz oder volle
Platte mitten im Schreiben zerstoert die Datei. Gleiches Muster wie die offene Memory-Notiz zum
Listen-Speichern: temp-Datei schreiben, dann ReplaceFileW bzw. rename. Bei Replace in Files ueber
hunderte Dateien ist das kein Nice-to-have.

### R3: Dateien ueber 2 GB
Encoding castet Laengen auf int (MultiByteToWideChar/WideCharToMultiByte). Ueber INT_MAX kippt das
in stilles Skippen oder Fehlverhalten; das Groessenlimit ist per Default aus. Minimum: Dateien
ueber einer harten Grenze sauber als SkipTooLarge behandeln statt in die Casts zu laufen.

### R4: Enumeration ohne Fortschritt und ohne Abbruch
Die Verzeichnis-Iteration prueft nur _isShuttingDown, kein _isCancelRequested, keine
Message-Pumpe, kein Fortschritt (ca. 9806). Bei einem grossen Baum friert die UI bis zum Ende der
Enumeration ein; ESC wirkt erst in der Dateischleife. N++ zeigt hier "Discovering file candidates"
mit Cancel. Message-Pumpe plus Cancel-Check in die Iterationsschleife, Statuszeile genuegt.

### R5: Speicherprofil
Pro Datei liegen zeitweise Rohbytes, UTF-8-Kopie und Scintilla-Kopie gleichzeitig im Speicher
(Faktor ~3). Fuer den Suchpfad durch den Binaerskip meist unkritisch, bei Replace in Files auf
grossen Textdateien relevant. Kein Handlungsbedarf fuer 6.1, aber bei Beschwerden der bekannte
Hebel (blockweises Verarbeiten).

### R6: BOM schaltet die Binaererkennung komplett ab
shouldSkipAsBinary: FF FE am Dateianfang macht jede Datei zu "Text", auch echte Binaerdateien mit
zufaelligem Prefix. Randfall; beim K1-Kontrakt mit abdecken (nach BOM-Erkennung Plausibilitaet der
Dekodierung pruefen, haengt mit R1-Roundtrip zusammen).

---

## 6. Geprueft und ausdruecklich OHNE Befund (Paritaet oder korrekt)

1. *.*-Default und PathMatchSpec: identische API und identischer Default wie N++, kein Delta im
   Files-Pfad (nur K2 intern).
2. Geklonte Dokumente in beiden Views werden doppelt gezaehlt: N++ iteriert genauso ueber beide
   DocTabs und zaehlt ebenfalls doppelt. Keine Abweichung, allenfalls spaeter Dedup per BufferID.
3. dotMatchesNL ist an allen 17 Aufrufstellen einheitlich false.
4. Suchbytes je Datei/Dokument mit korrekter Codepage erzeugt (Binding vor Kontextaufbau).
5. Delimiter-Verarbeitung kostet im Nicht-Spaltenmodus nichts (early return,
   handleDelimiterPositions ca. 13969).
6. Skip-Zaehler starten pro Operation bei 0 (Guard lokal, create() setzt zurueck).
7. Dock-Anzeige gegen Riesenzeilen gekappt (2048 Bytes je Treffer).

---

## 7. Verifikationsplan

1. Eigener Testkorpus, unabhaengig von guy038: identischer Text in ANSI (1252), UTF-8,
   UTF-8-BOM, UTF-16LE-BOM, UTF-16BE-BOM, UTF-16LE-ohne-BOM, UTF-16BE-ohne-BOM; dazu eine echte
   Binaerdatei, eine endungslose Datei, eine versteckte Datei, eine Datei in verstecktem Ordner,
   eine UTF-16-Datei mit ungerader Bytezahl. Erwartete Trefferzahlen je Tool (N++, MultiReplace
   vor/nach Fix) als Tabelle festhalten; Generator-Skript folgt mit der Umsetzung.
2. guy038s E:\Test-Ordner (Issue #133) gegen dieselbe Matrix laufen lassen; erst danach die
   Bewertung im Forum abschliessen.
3. K1-Logik (Erkennung/Konvertierung) als Header-only-Einheit im Container mit
   g++ -fsanitize=address,undefined testen: Roundtrips fuer alle Encodings, Grenzfaelle
   (ungerade Laenge, lone surrogates, ungueltige Bytes, 0-Byte-Datei, nur-BOM-Datei).
4. Referenzmessung gegen N++ auf einem gemischten Verzeichnis (Zeit, Treffer, Dateien,
   Skip-Zaehler) als Regressionsanker vor jedem Release.

---

## 8. Empfohlene Reihenfolge fuer 6.1

1. F2 zuerst: isolierter Einzeiler, Crash-Klasse, kein Risiko.
2. K1-Pipeline bauen und Find/Replace in Files darauf umstellen; F1 und F3 fallen dabei als
   Konsequenz des Kontrakts heraus, R1-Roundtrip und R6 als Hooks gleich mit anlegen.
3. F4 Transparenz: Skip-Zaehler je Grund, Header "X hits in Y of M files", Statuszeile;
   gemeinsam mit der Binary-Option (an/aus in den Settings neben dem Groessenlimit) als ein
   Settings-Block. Das ist der bereits geplante naechste Schritt.
4. K3 und K2 entscheiden und umsetzen (Hidden-Semantik, eine Filter-Engine).
5. R2 atomares Schreiben, R4 Enumeration mit Cancel, R3 2-GB-Grenze.
6. K5 Zero-Length-Zaehlung angleichen.
7. README: Kapitel "File Search" mit Skip-Regeln, Filter- und Hidden-Semantik dokumentieren.

Punkt 1 ist unabhaengig sofort machbar. Punkte 2 und 3 bilden zusammen die Antwort auf Issue #133.

---

## 9. Umsetzungsstand (2026-08-30)

Umgesetzt in einem Zug, Quelldateien im Repo aktualisiert:

| Punkt | Umsetzung |
|---|---|
| F1/F3/K1 | Gemeinsame Pipeline HiddenSciGuard::loadTextFile (Header, BOM, UTF-16-Erkennung, Binaercheck, Volllesen, Dekodieren); isLikelyBinary entfernt; Find UND Replace in Files nutzen dieselbe Pipeline. UTF-16 LE ohne BOM per IsTextUnicode wie N++ (Encoding::detectEncoding, Option enableUtf16NoBomLE); BE ohne BOM bewusst nicht (N++-Paritaet, kommentiert) |
| F2 | Docs-collect setzt findTextW aus dem Parameter, Single Mode uebergibt findW; OOB beseitigt |
| F4 | Vier Skip-Zaehler im Guard (binary/too large/unreadable/undecodable), Dock-Header-Suffix "[of M searched, N skipped: ...]" via ResultDock::closeSearchBlock(suffix) NACH dem Platzhalter-Patch (Scan bleibt stabil); Replace-Statuszeile mit Suffix + read-only-Zaehler |
| Option | Settings "File Search: Skip binary files" (Default an), IDC_CFG_SKIP_BINARY, Binding, INI [ReplaceInFiles] SkipBinaryFiles, Defaults; Skip aus = RawBytes-Suche wie N++, Replace schreibt Rohbytes verbatim zurueck |
| K3 | collectScanFiles: versteckte ORDNER werden bei ausgeschalteter Checkbox geprunt (disable_recursion_pending), versteckte Dateien immer dabei; matchPath nur noch Pattern (N++-Semantik, passend zum Label "In hidden folders") |
| K5 | h.length>0-Filter in allen vier Sammelpfaden entfernt; Zero-Length-Treffer zaehlen wie bei N++ |
| R1 | setVerifyRoundtrip(true) nur im Replace-Pfad: Encoding::verifyLosslessDecode (Re-Encode + memcmp) lehnt verlustbehaftete Dekodierungen als Undecodable ab |
| R2 | writeFile atomar: .mr_tmp + ReplaceFileW, Fallback MoveFileExW, Cleanup bei Fehler |
| R3 | INT_MAX-Guard in convertBufferToUtf8 -> Undecodable statt UB |
| R4 | Enumeration mit Message-Pumpe, ESC-Cancel, Statusanzeige (status_discovering_files), Fehlerpfad erhalten |
| K2 | Beide Files-Pfade nutzen weiterhin PathMatchSpecW wie N++ (Paritaet per Konstruktion); matchesDocFilter (Docs) unveraendert, endungslose Frage wird per Korpus geklaert |
| Doku | README: Option, Scan-Statistik, Encoding-/Hidden-Semantik; languages.ini EN+DE plus language_mapping.cpp (9 neue Keys) |

| F5/F6 | fileScope-Diskriminator (explicitPath) in handleReplaceAllButton und replaceAll (neuer Parameter, Default false): allFromCursor und isSelectionMode werden im Datei-Scope maskiert; zusaetzlich SCI_GOTOPOS 0 nach dem Laden als definierter Buffer-Zustand |
| Settings | Eigene Kategorie "File Search" (Position 2): neues Panel (_hFileSearchPanel, IDC_CFG_GRP_FILE_SEARCH 7909), Gruppe "File handling" mit Groessenlimit + Binaer-Checkbox, Bindings/Labels/DarkMode/MoveWindow/panels[]/ShowWindow-Indizes symmetrisch erweitert; "Search behaviour" zurueck auf 160 px; Labels ohne "File Search:"-Praefix (EN/DE), Kategorie- und Gruppen-Keys neu; README-Abschnitt "2. File Search", Folgeabschnitte umnummeriert |

Container-QS: src/tests/file_search_qa.cpp neu geschrieben auf Post-Fix-Logik, 11/11 PASS unter
ASan/UBSan (UTF-16-BOM dekodiert, LE-ohne-BOM erkannt, BE-ohne-BOM Paritaets-Skip, odd-length als
Undecodable gezaehlt, Skip-an/aus-Matrix, NUL-jenseits-8KB als Text, gesunde Pfade unveraendert).
Brace/Paren-Balance aller editierten Dateien gegen Baseline verifiziert, BOM/CRLF je Datei erhalten.

## 10. Windows-QS-Checkliste (vor Release)

1. Build in VS, 0 error-Zeilen.
2. src\tests\gen_search_corpus.ps1 ausfuehren; Suche "Fi" (Regex, Match case, *.*) gegen die im
   Skript ausgegebene Erwartungstabelle: Skip an, Skip aus, Hidden an/aus; noext-Zeile notieren
   (klaert K2/PathMatchSpec).
3. Gegenprobe N++ nativ auf demselben Korpus: identische Dateimenge bei Skip aus.
4. Zero-Length: Regex ^ auf einer Datei mit N Zeilen -> N Treffer, identisch zu N++.
5. Replace in Files auf dem Korpus (harmlose Ersetzung): utf16le/be-BOM und LE-ohne-BOM werden
   korrekt zurueckgeschrieben (BOM und Byteorder erhalten), utf16-odd wird als not decodable
   uebersprungen, Rohdatei bleibt unangetastet; .mr_tmp-Dateien duerfen nie zurueckbleiben.
6. Settings-Panel: neue Checkbox sichtbar, Persistenz nach Neustart, Default an.
7. Dock-Header: Suffix erscheint nur bei Find in Files, Klick-Navigation auf Treffer nach Suffix
   weiterhin korrekt (Offset-Verschiebung), Sprachumschaltung EN/DE.
8. guy038s E:\Test-Ordner (sobald im Issue #133): Zahlen gegen N++ und gegen Erwartung pruefen.
9. F5-Regression: "Find: Search from cursor position" AN, Wrap aus, Replace in Files auf
   cursortest.txt (3x Alpha -> BETA): jetzt 3 Ersetzungen (vorher 0). Danach Option AUS
   wiederholen: unveraendert 3. Editor-Verhalten unveraendert: im offenen Dokument startet
   Replace All mit Option AN weiterhin ab Cursor.
10. F6-Regression: Selection-Radio aktiv lassen und Replace in Files starten: es wird die ganze
    Datei ersetzt, nicht nur ein Selektionsbereich; anschliessend im Editor pruefen, dass
    Selection-Replace dort weiterhin nur die Auswahl trifft.
11. Settings-Kategorie: "File Search" an Position 2, Panel zeigt Gruppe "File handling" mit
    beiden Optionen (Labels ohne Praefix), MB-Feld koppelt an die Checkbox, Kategorienwechsel
    in alle Richtungen, DarkMode, Sprachumschaltung EN/DE, "Reset All Settings" setzt Default
    (Limit aus, Binaer-Skip an).
