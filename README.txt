NEPLAN-Projektdatei exportieren
===============================

Macht aus NEPLAN-Projektdateien (.nepprj) Access-Datenbanken, wie sie der
Datenbankexport von NEPLAN liefert: dieselben 31 Tabellen mit allen Spalten.
NEPLAN wird dafür nicht gebraucht, auch kein Python, nur Windows.


So geht es
----------
1. Doppelklick auf "NEPLAN-Projektdatei exportieren".
2. "Projektdateien auswählen …" klicken und eine oder mehrere .nepprj wählen
   (mit Strg oder Umschalt). Hineinziehen ins Fenster geht auch.
3. "Exportieren" klicken.

Je Projektdatei entsteht eine Datenbank <Ort>.mdb in C:\Users\<name>\NEPLAN-Export,
zum Beispiel Dingden.mdb. Gibt es sie schon, wird sie ersetzt. Über
"Ordner öffnen" kommt man direkt hin. Diese Datenbanken nimmt der
Schaltplan-Marker im Reiter "Netzabgleich".

Die Projektdateien werden nur gelesen, auch vom gemeinsamen Laufwerk. Sie
dürfen dabei in NEPLAN offen sein.


Was in der Datenbank steht
--------------------------
Alles, was auch der NEPLAN-Export enthält: Stationen, Leitungen mit ihren
Kabelabschnitten, Trenn-, Last- und Leistungsschalter, Kupplung, Trafo,
Einspeisung, Lasten, Erzeuger und Relais mit allen elektrischen Werten, die
Topologie mit Schaltzustand, Abgang, Teilnetz und Zone, dazu Farben,
Skalierungsfaktoren und Benutzerdatenfelder. Tabellen für Elementarten, die im
Netz nicht vorkommen, bleiben wie bei NEPLAN leer.
Zusätzlich zum NEPLAN-Export steht die Änderungshistorie aus dem Schriftfeld
des Plans in der Tabelle A_AENDERUNGEN.

Geprüft am 11.09.2026 an Dingden gegen den NEPLAN-Export derselben Datei: alle
95.212 Werte gleich. Wo die Werte in der Projektdatei stehen, wurde an Dingden
ermittelt. Bei einem neuen Ort daher einmal prüfen (siehe unten), bevor man
sich darauf verlässt.


Prüfen an anderen Orten
-----------------------
Im Unterordner "Prüfen": eine Datenbank, die NEPLAN selbst exportiert hat, auf
"Leser gegen NEPLAN prüfen" ziehen. Das Werkzeug exportiert die passende
Projektdatei selbst und vergleicht beide Datenbanken Wert für Wert über alle
Tabellen. Steht am Ende "stimmt vollständig", passt alles. Sonst zeigt der
Bericht je Tabelle, welche Spalten abweichen. Die Projektdatei muss dafür
denselben Stand haben wie der NEPLAN-Export. Bericht und selbst erzeugte
Datenbank landen in C:\Users\<name>\NEPLAN-Export\Berichte.


Aktualisierung
--------------
Unten rechts im Fenster stehen die Version und der Knopf "Nach Updates suchen".
Beim Start sieht das Werkzeug selbst nach, ob es eine neuere Version gibt. Wenn
ja, wird der Knopf blau ("Update auf …"). Ein Klick zeigt, was neu ist, und
aktualisiert nach Rückfrage; das Fenster startet danach neu. Geladen wird von
GitHub. Geht dabei etwas schief, bleibt die bisherige Version unverändert.
Exportierte Datenbanken und Einstellungen werden nicht angefasst.


Dateien
-------
NEPLAN-Projektdatei exportieren.cmd   der Starter
Prüfen\                               Vergleich mit einem echten NEPLAN-Export
Vorlage\                              leere Datenbank mit allen Tabellen des
                                      NEPLAN-Exports, aus der jede neue entsteht
werkzeug\                             Fenster und Aktualisierung (PowerShell),
                                      Leser (C#, wird von Windows beim Start
                                      übersetzt)
manifest.json                         Version und Prüfsummen für die Aktualisierung
