// Liest eine NEPLAN-Projektdatei (.nepprj) ohne NEPLAN.
//
// Die Datei ist ein OLE-Container, das Netzmodell steht im Datenstrom "Root/Data"
// als MFC-Serialisierung. Entschluesselt am 11.09.2026 an Dingden, geprueft gegen
// den NEPLAN-Datenbankexport derselben Datei:
//   Texte      FF FE FF <Laenge> <UTF-16>
//   Element    int32 NEPID | Text NAME | Text TYP        (direkt hintereinander)
//   Felder     vor dem Namen: ALIASNAME 2 Texte davor, DESCRIPTION 9 (Station) bzw.
//              5 Texte davor, legacy/panelname hinter ihrer Beschriftung
//   Anschluss  int32 2 | Element | Station | Klemme | Zustand   (oder Station | Element)
// Geschrieben fuer den C#-5-Compiler von Windows PowerShell 5.1 (Add-Type).

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace NeplanLeser
{
    // ------------------------------------------------------------------ //
    // OLE-Container (Compound File Binary)
    // ------------------------------------------------------------------ //
    public class Cfb
    {
        private readonly byte[] f;
        private readonly int secSize, miniSecSize;
        private readonly uint miniCutoff;
        private readonly List<int> fat = new List<int>();
        private readonly List<int> miniFat = new List<int>();
        private readonly List<Eintrag> dir = new List<Eintrag>();
        private readonly byte[] miniStream;

        private class Eintrag
        {
            public string Name;
            public int Links, Rechts, Kind, Start;
            public long Groesse;
        }

        public Cfb(byte[] daten)
        {
            f = daten;
            if (f.Length < 512 || BitConverter.ToUInt64(f, 0) != 0xE11AB1A1E011CFD0UL)
                throw new InvalidDataException("Das ist keine NEPLAN-Projektdatei (kein OLE-Container).");
            secSize = 1 << BitConverter.ToUInt16(f, 0x1E);
            miniSecSize = 1 << BitConverter.ToUInt16(f, 0x20);
            int anzFat = BitConverter.ToInt32(f, 0x2C);
            int ersterDir = BitConverter.ToInt32(f, 0x30);
            miniCutoff = BitConverter.ToUInt32(f, 0x38);
            int ersteMiniFat = BitConverter.ToInt32(f, 0x3C);
            int ersteDifat = BitConverter.ToInt32(f, 0x44);
            int anzDifat = BitConverter.ToInt32(f, 0x48);

            var difat = new List<int>();
            for (int i = 0; i < 109; i++) difat.Add(BitConverter.ToInt32(f, 0x4C + 4 * i));
            int s = ersteDifat;
            for (int n = 0; n < anzDifat && s >= 0; n++)
            {
                int off = Offset(s);
                int proSektor = secSize / 4 - 1;
                for (int i = 0; i < proSektor; i++) difat.Add(BitConverter.ToInt32(f, off + 4 * i));
                s = BitConverter.ToInt32(f, off + 4 * proSektor);
            }
            for (int i = 0; i < anzFat && i < difat.Count; i++)
            {
                int off = Offset(difat[i]);
                for (int k = 0; k < secSize / 4 && off + 4 * k + 4 <= f.Length; k++)
                    fat.Add(BitConverter.ToInt32(f, off + 4 * k));
            }

            byte[] d = Kette(ersterDir, fat, secSize, null, -1);
            for (int off = 0; off + 128 <= d.Length; off += 128)
            {
                var e = new Eintrag();
                int len = BitConverter.ToUInt16(d, off + 0x40);
                e.Name = len >= 2 && len <= 64 ? Encoding.Unicode.GetString(d, off, len - 2) : "";
                e.Links = BitConverter.ToInt32(d, off + 0x44);
                e.Rechts = BitConverter.ToInt32(d, off + 0x48);
                e.Kind = BitConverter.ToInt32(d, off + 0x4C);
                e.Start = BitConverter.ToInt32(d, off + 0x74);
                e.Groesse = BitConverter.ToUInt32(d, off + 0x78);
                dir.Add(e);
            }
            if (ersteMiniFat >= 0)
            {
                byte[] mf = Kette(ersteMiniFat, fat, secSize, null, -1);
                for (int k = 0; k + 4 <= mf.Length; k += 4) miniFat.Add(BitConverter.ToInt32(mf, k));
            }
            miniStream = dir.Count > 0 ? Kette(dir[0].Start, fat, secSize, null, dir[0].Groesse) : new byte[0];
        }

        private int Offset(int sektor) { return (sektor + 1) * secSize; }

        // Folgt einer Sektorkette. quelle == null heisst Datei ueber die FAT, sonst Mini-Stream.
        private byte[] Kette(int start, List<int> tabelle, int groesse, byte[] quelle, long laenge)
        {
            var ms = new MemoryStream();
            byte[] src = quelle ?? f;
            int s = start, schritte = 0;
            while (s >= 0 && s < tabelle.Count && schritte++ <= tabelle.Count)
            {
                int off = quelle == null ? Offset(s) : s * groesse;
                int n = Math.Min(groesse, src.Length - off);
                if (n <= 0) break;
                ms.Write(src, off, n);
                s = tabelle[s];
            }
            byte[] r = ms.ToArray();
            if (laenge >= 0 && laenge < r.Length) Array.Resize(ref r, (int)laenge);
            return r;
        }

        public byte[] LiesStrom(string pfad)
        {
            int akt = 0;
            foreach (string teil in pfad.Split('/'))
            {
                int gefunden = Suche(dir[akt].Kind, teil);
                if (gefunden < 0) throw new InvalidDataException("In der Projektdatei fehlt der Datenstrom " + pfad + ".");
                akt = gefunden;
            }
            Eintrag e = dir[akt];
            if (e.Groesse < miniCutoff) return Kette(e.Start, miniFat, miniSecSize, miniStream, e.Groesse);
            return Kette(e.Start, fat, secSize, null, e.Groesse);
        }

        private int Suche(int wurzel, string name)
        {
            var stapel = new Stack<int>();
            stapel.Push(wurzel);
            int schritte = 0;
            while (stapel.Count > 0 && schritte++ <= dir.Count * 2)
            {
                int i = stapel.Pop();
                if (i < 0 || i >= dir.Count) continue;
                if (string.Equals(dir[i].Name, name, StringComparison.OrdinalIgnoreCase)) return i;
                stapel.Push(dir[i].Links);
                stapel.Push(dir[i].Rechts);
            }
            return -1;
        }
    }

    // ------------------------------------------------------------------ //
    // Ergebnis
    // ------------------------------------------------------------------ //
    public class Text
    {
        public int Start, Ende;
        public string Wert;
    }

    public class Element
    {
        public int Nepid;
        public string Name = "", Typ = "", Alias = "", Beschreibung = "", Legacy = "", Panelname = "";
        public int NameIndex;
        public int[] Knoten = { 0, 0, 0, 0 };        // NEPID der Station je Klemme, 0 = keine
        public int[] Zustand = { -1, -1, -1, -1 };   // 1 = geschlossen, 0 = offen, -1 = unbekannt
        public int Saetze;
        public bool Widerspruch;
        public int Leitung;        // Schalter am Leitungsende: NEPID der Leitung, sonst 0
        public int NepidPos;       // Lage der NEPID im Datenstrom
    }

    public class Aenderung
    {
        public string Rev, Nummer, Datum, Kuerzel;
    }

    public class Projekt
    {
        public string Datei, Planname = "";
        public Dictionary<int, Element> Elemente = new Dictionary<int, Element>();
        public List<Element> Reihenfolge = new List<Element>();
        public List<Aenderung> Aenderungen = new List<Aenderung>();
        public List<string> Hinweise = new List<string>();

        public Element Station(int nepid)
        {
            Element e;
            return nepid != 0 && Elemente.TryGetValue(nepid, out e) ? e : null;
        }
    }

    // ------------------------------------------------------------------ //
    // Leser
    // ------------------------------------------------------------------ //
    public static class Leser
    {
        public const string Version = "nepprj-Leser 1.0";

        // Elementarten, die auch ohne Anschlusssatz uebernommen werden
        private static readonly HashSet<string> Bekannt = new HashSet<string> {
            "BUSBAR-NODE", "LINE", "LOAD", "AC_GENERIC_COMP", "DISCSWITCH_2", "LOADSWITCH",
            "CIRC_BREAKER_2", "TRANSFORMER", "FEEDER", "COUPLING", "OVERCUR_RELAIS", "VOL_RELAIS" };
        // Beschreibung steht 5 Texte vor dem Namen; nur fuer diese Arten geprueft
        private static readonly HashSet<string> MitBeschreibung = new HashSet<string> {
            "LOAD", "DISCSWITCH_2", "LOADSWITCH", "CIRC_BREAKER_2", "COUPLING" };
        private static readonly Regex TypMuster = new Regex(@"^[A-Z][A-Z0-9_\-]{2,}$");
        private static readonly Regex DatumMuster = new Regex(@"^\d{1,2}\.\d{1,2}\.\d{2,4}$");

        public static Projekt Lies(string datei)
        {
            var p = new Projekt();
            p.Datei = Path.GetFullPath(datei);
            byte[] d = new Cfb(LiesDatei(p.Datei)).LiesStrom("Root/Data");
            List<Text> cs = Texte(d);

            var typEnden = new List<int>();
            FindeElemente(p, d, cs, typEnden);
            LiesFelder(p, cs);
            Dictionary<int, int[]> klemmen = FindeAnschluesse(p, d);
            Aufraeumen(p);
            AmLeitungsende(p, d, typEnden, klemmen);
            LiesAenderungen(p, cs);
            return p;
        }

        // Liest die Datei auch, wenn sie gerade in NEPLAN offen ist. Aendert sie sich
        // waehrend des Lesens (etwa beim woechentlichen Hochladen), wird neu gelesen.
        public static byte[] LiesDatei(string pfad)
        {
            for (int versuch = 1; ; versuch++)
            {
                var vorher = new FileInfo(pfad);
                long laenge = vorher.Length;
                DateTime zeit = vorher.LastWriteTimeUtc;
                byte[] daten;
                using (var fs = new FileStream(pfad, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                {
                    daten = new byte[fs.Length];
                    int gelesen = 0;
                    while (gelesen < daten.Length)
                    {
                        int n = fs.Read(daten, gelesen, daten.Length - gelesen);
                        if (n <= 0) break;
                        gelesen += n;
                    }
                    if (gelesen < daten.Length) Array.Resize(ref daten, gelesen);
                }
                var nachher = new FileInfo(pfad);
                if (nachher.Length == laenge && nachher.LastWriteTimeUtc == zeit) return daten;
                if (versuch >= 3)
                    throw new IOException("Die Projektdatei ändert sich gerade, wird sie hochgeladen? Bitte später noch einmal versuchen.");
                System.Threading.Thread.Sleep(2000);
            }
        }

        public static List<Text> Texte(byte[] d)
        {
            var r = new List<Text>();
            for (int i = 0; i + 4 <= d.Length; i++)
            {
                if (d[i] != 0xFF || d[i + 1] != 0xFE || d[i + 2] != 0xFF) continue;
                int p = i + 3, n = d[p++];
                if (n == 0xFF)
                {
                    if (p + 2 > d.Length) break;
                    n = BitConverter.ToUInt16(d, p); p += 2;
                    if (n == 0xFFFF)
                    {
                        if (p + 4 > d.Length) break;
                        n = BitConverter.ToInt32(d, p); p += 4;
                    }
                }
                int naechster = i + 2;
                if (n >= 0 && n <= 20000 && p + 2 * n <= d.Length)
                {
                    string w = Encoding.Unicode.GetString(d, p, 2 * n);
                    if (!KaputteSurrogate(w)) r.Add(new Text { Start = i, Ende = p + 2 * n, Wert = w });
                }
                i = naechster;
            }
            return r;
        }

        private static bool KaputteSurrogate(string s)
        {
            for (int i = 0; i < s.Length; i++)
            {
                if (char.IsHighSurrogate(s[i]))
                {
                    if (i + 1 >= s.Length || !char.IsLowSurrogate(s[i + 1])) return true;
                    i++;
                }
                else if (char.IsLowSurrogate(s[i])) return true;
            }
            return false;
        }

        // typEnden sammelt das Ende jedes Objektkopfs, auch der Klemmen (EDGE); daraus
        // ergibt sich, wo das Objekt vor einer NEPID beginnt.
        private static void FindeElemente(Projekt p, byte[] d, List<Text> cs, List<int> typEnden)
        {
            var nachEnde = new Dictionary<int, int>();
            for (int i = 0; i < cs.Count; i++) nachEnde[cs[i].Ende] = i;
            for (int i = 0; i < cs.Count; i++)
            {
                string typ = cs[i].Wert;
                if (!TypMuster.IsMatch(typ)) continue;
                int j;
                if (!nachEnde.TryGetValue(cs[i].Start, out j)) continue;
                Text nm = cs[j];
                if (nm.Wert.Length == 0 || nm.Start < 4) continue;
                typEnden.Add(cs[i].Ende);
                if (typ == "EDGE") continue;
                int nep = BitConverter.ToInt32(d, nm.Start - 4);
                if (nep <= 0 || p.Elemente.ContainsKey(nep)) continue;
                var e = new Element { Nepid = nep, Name = nm.Wert, Typ = typ, NameIndex = j, NepidPos = nm.Start - 4 };
                p.Elemente[nep] = e;
                p.Reihenfolge.Add(e);
            }
        }

        private static string Hole(List<Text> cs, int i)
        {
            return i >= 0 && i < cs.Count ? cs[i].Wert : "";
        }

        private static void LiesFelder(Projekt p, List<Text> cs)
        {
            int untergrenze = 0;
            foreach (Element e in p.Reihenfolge)
            {
                int j = e.NameIndex;
                e.Alias = Hole(cs, j - 2);
                if (e.Typ == "BUSBAR-NODE") e.Beschreibung = Hole(cs, j - 9);
                else if (MitBeschreibung.Contains(e.Typ)) e.Beschreibung = Hole(cs, j - 5);
                for (int x = j - 1; x > untergrenze && x >= j - 14; x--)
                {
                    if (cs[x].Wert == "legacy" && e.Legacy.Length == 0) e.Legacy = Hole(cs, x + 1);
                    if (cs[x].Wert == "panelname" && e.Panelname.Length == 0) e.Panelname = Hole(cs, x + 1);
                }
                untergrenze = j + 1;
            }
        }

        // Liefert zusaetzlich die Klemmennummern der Leitungsenden: Klemme -> {Leitung, Klemmenindex}.
        // Die Klemmennummer steht 64 Bytes hinter dem ersten ID-Feld des Anschlusssatzes.
        private static Dictionary<int, int[]> FindeAnschluesse(Projekt p, byte[] d)
        {
            var klemmen = new Dictionary<int, int[]>();
            var stationen = new HashSet<int>();
            var andere = new HashSet<int>();
            foreach (Element e in p.Reihenfolge)
                (e.Typ == "BUSBAR-NODE" ? stationen : andere).Add(e.Nepid);

            for (int off = 4; off + 16 <= d.Length; off++)
            {
                if (BitConverter.ToInt32(d, off - 4) != 2) continue;
                int kl = BitConverter.ToInt32(d, off + 8);
                int z = BitConverter.ToInt32(d, off + 12);
                if (kl < 0 || kl > 3 || (z != 0 && z != 1)) continue;
                int a = BitConverter.ToInt32(d, off), b = BitConverter.ToInt32(d, off + 4);
                int el, st;
                if (andere.Contains(a) && stationen.Contains(b)) { el = a; st = b; }
                else if (stationen.Contains(a) && andere.Contains(b)) { el = b; st = a; }
                else continue;
                Element e = p.Elemente[el];
                e.Saetze++;
                if (e.Knoten[kl] == 0) { e.Knoten[kl] = st; e.Zustand[kl] = z; }
                else if (e.Knoten[kl] != st || e.Zustand[kl] != z) e.Widerspruch = true;
                if (e.Typ == "LINE" && off + 68 <= d.Length)
                {
                    int kid = BitConverter.ToInt32(d, off + 64);
                    if (!klemmen.ContainsKey(kid)) klemmen[kid] = new[] { el, kl };
                }
            }
            return klemmen;
        }

        // Unbekannte Arten ohne Anschluss sind Rechenparameter o. Ae. und fliegen raus.
        private static void Aufraeumen(Projekt p)
        {
            var neu = new List<Element>();
            var unbekannt = new Dictionary<string, int>();
            foreach (Element e in p.Reihenfolge)
            {
                if (Bekannt.Contains(e.Typ)) { neu.Add(e); continue; }
                if (e.Saetze > 0)
                {
                    neu.Add(e);
                    int n; unbekannt.TryGetValue(e.Typ, out n); unbekannt[e.Typ] = n + 1;
                    continue;
                }
                p.Elemente.Remove(e.Nepid);
            }
            p.Reihenfolge = neu;
            foreach (var kv in unbekannt)
                p.Hinweise.Add("Unbekannte Elementart " + kv.Key + " (" + kv.Value + " Stück) mit Anschlüssen übernommen.");
            int wid = 0;
            foreach (Element e in p.Reihenfolge) if (e.Widerspruch) wid++;
            if (wid > 0) p.Hinweise.Add(wid + " Elemente mit widersprüchlichen Anschlusssätzen, jeweils der erste wurde genommen.");
        }

        // Trennschalter, Leistungsschalter und Schutz sitzen an einem Leitungsende und haben
        // keinen eigenen Anschlusssatz. In ihrem Objekt, vor der NEPID, steht genau eine
        // Klemmennummer eines Leitungsendes. Daraus folgen Leitung, Station und Zustand.
        // Beim Trennschalter steht die Station zusaetzlich im ALIASNAME, das dient als Kontrolle.
        private static readonly HashSet<string> AnLeitung = new HashSet<string> { "DISCSWITCH_2", "CIRC_BREAKER_2", "OVERCUR_RELAIS" };

        private static void AmLeitungsende(Projekt p, byte[] d, List<int> typEnden, Dictionary<int, int[]> klemmen)
        {
            var stationNachName = new Dictionary<string, int>();
            foreach (Element e in p.Reihenfolge)
                if (e.Typ == "BUSBAR-NODE" && !stationNachName.ContainsKey(e.Name.Trim())) stationNachName[e.Name.Trim()] = e.Nepid;
            int ohneLeitung = 0, widerspruch = 0;
            foreach (Element e in p.Reihenfolge)
            {
                if (!AnLeitung.Contains(e.Typ) || e.Saetze > 0) continue;
                int idx = typEnden.BinarySearch(e.NepidPos);
                if (idx < 0) idx = ~idx;
                int von = idx > 0 ? typEnden[idx - 1] : 0;
                var gefunden = new HashSet<int>();
                for (int q = von; q + 4 <= e.NepidPos; q++)
                {
                    int v = BitConverter.ToInt32(d, q);
                    if (klemmen.ContainsKey(v)) gefunden.Add(v);
                }
                int stAlias = 0;
                bool hatAlias = e.Typ == "DISCSWITCH_2" && stationNachName.TryGetValue(e.Alias.Trim(), out stAlias);
                if (gefunden.Count == 1)
                {
                    int kid = 0;
                    foreach (int x in gefunden) kid = x;
                    int[] ende = klemmen[kid];
                    Element l = p.Elemente[ende[0]];
                    e.Leitung = l.Nepid;
                    e.Knoten[1] = l.Knoten[ende[1]];
                    e.Zustand[0] = l.Zustand[ende[1]];
                    e.Zustand[1] = l.Zustand[ende[1]];
                    if (hatAlias && stAlias != e.Knoten[1]) widerspruch++;
                }
                else
                {
                    if (hatAlias) e.Knoten[1] = stAlias;    // Station bekannt, Leitung und Zustand nicht
                    if (e.Typ == "DISCSWITCH_2") ohneLeitung++;
                }
            }
            if (ohneLeitung > 0) p.Hinweise.Add(ohneLeitung + " Trennschalter ohne eindeutige Leitung, dort ist nur die Station eingetragen.");
            if (widerspruch > 0) p.Hinweise.Add(widerspruch + " Trennschalter, bei denen Leitungsende und ALIASNAME verschiedene Stationen nennen.");
        }

        // Aenderungshistorie aus dem Schriftfeld: hinter "Rev." der Kopf "Aenderungen | Datum | Name",
        // dann je Eintrag Revisionsbuchstabe | Nummer | Datum | Kuerzel. Der Planname steht hinter
        // dem naechsten "$PROJNAME".
        private static void LiesAenderungen(Projekt p, List<Text> cs)
        {
            for (int i = 0; i < cs.Count; i++)
            {
                if (cs[i].Wert != "Rev.") continue;
                var werte = new List<string>();
                int k = i + 1;
                for (; k < cs.Count && werte.Count < 400; k++)
                {
                    string w = cs[k].Wert.Trim();
                    if (w.Length == 0) continue;
                    werte.Add(w);
                    if (w == "$PROJNAME") break;
                }
                int x = werte.IndexOf("Name") + 1;
                while (x + 3 < werte.Count && werte[x].Length <= 3 && DatumMuster.IsMatch(werte[x + 2]))
                {
                    p.Aenderungen.Add(new Aenderung { Rev = werte[x], Nummer = werte[x + 1], Datum = werte[x + 2], Kuerzel = werte[x + 3] });
                    x += 4;
                }
                if (p.Aenderungen.Count == 0)
                {
                    // Plaene ohne Revisionsbuchstaben: Nummer | Datum | Kuerzel
                    x = werte.IndexOf("Name") + 1;
                    while (x + 2 < werte.Count && DatumMuster.IsMatch(werte[x + 1]))
                    {
                        p.Aenderungen.Add(new Aenderung { Rev = "", Nummer = werte[x], Datum = werte[x + 1], Kuerzel = werte[x + 2] });
                        x += 3;
                    }
                }
                if (werte.Count > 0 && werte[werte.Count - 1] == "$PROJNAME")
                    for (int m = k + 1; m < cs.Count && m < k + 6; m++)
                        if (cs[m].Wert.Trim().Length > 0) { p.Planname = cs[m].Wert.Trim(); break; }
                break;
            }
        }

        // Letzte Aenderung als Datum, falls lesbar
        public static DateTime? LetzteAenderung(Projekt p)
        {
            DateTime? best = null;
            string[] formate = { "d.M.yyyy", "d.M.yy", "dd.MM.yyyy", "dd.MM.yy" };
            foreach (Aenderung a in p.Aenderungen)
            {
                DateTime t;
                if (DateTime.TryParseExact(a.Datum, formate, System.Globalization.CultureInfo.GetCultureInfo("de-DE"),
                        System.Globalization.DateTimeStyles.None, out t) && (best == null || t > best)) best = t;
            }
            return best;
        }
    }
}
