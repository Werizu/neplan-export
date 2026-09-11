// Vollexport einer NEPLAN-Projektdatei in eine Access-Datenbank mit allen Tabellen und
// Spalten des NEPLAN-Datenbankexports, ohne NEPLAN.
//
// Nachbau von vollexport_proto.py. Am 11.09.2026 an Dingden Zelle fuer Zelle gegen den
// NEPLAN-Export derselben Datei geprueft: 95.212 von 95.212 Zellen gleich.
// Die Feldpositionen stehen in Regeln.cs (gelernt), hier stehen Topologie, die von NEPLAN
// beim Export berechneten Spalten (Abgang, Teilnetz, Zone), Farbtabelle, Faktoren,
// Benutzerdaten und das Schreiben. Dazu der Zellvergleich fuer das Pruefwerkzeug.
// Geschrieben fuer den C#-5-Compiler von Windows PowerShell 5.1 (Add-Type), ohne LINQ.

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace NeplanLeser
{
    public class Regel
    {
        public readonly string Tabelle, Spalte, Art;
        public readonly int K, Abstand;
        public readonly char Format;

        public Regel(string tabelle, string spalte, string art) { Tabelle = tabelle; Spalte = spalte; Art = art; }
        public Regel(string tabelle, string spalte, string art, int k) : this(tabelle, spalte, art) { K = k; }
        public Regel(string tabelle, string spalte, string art, int k, int abstand, char format) : this(tabelle, spalte, art)
        {
            K = k; Abstand = abstand; Format = format;
        }
    }

    public class Objekt
    {
        public int Nepid, A, E, NameIdx, TypIdx, NamePos, Anker;
        public string Typ, Name;
        public int[] Texte;                         // Indizes der Texte im Objekt
        public int[] Knoten = { 0, 0, 0, 0 };       // NEPID der Station je Klemme
        public int[] Zustand = { -1, -1, -1, -1 };
    }

    public class Vollexport
    {
        public const string Version = "nepprj-Leser 2.0 (ohne NEPLAN)";

        public static readonly Dictionary<string, string> Tabelle = new Dictionary<string, string> {
            { "BUSBAR-NODE", "BUSBAR" }, { "LINE", "LINE" }, { "LOAD", "LOAD" }, { "AC_GENERIC_COMP", "AC_GENERIC_COMP" },
            { "DISCSWITCH_2", "DISC_SWITCH" }, { "LOADSWITCH", "LOAD_SWITCH" }, { "CIRC_BREAKER_2", "CIRC_BREAK" },
            { "TRANSFORMER", "TRAFO2" }, { "FEEDER", "NETWFEEDER" }, { "COUPLING", "BB_COUPLER" },
            { "OVERCUR_RELAIS", "RELAY_OVERCUR" }, { "VOL_RELAIS", "RELAY_MINMAX_ON_NODE" } };
        private static readonly Regex TypMuster = new Regex(@"^[A-Z][A-Z0-9_\-]{2,}$");
        private static readonly HashSet<string> Schaltgeraet = new HashSet<string> { "LOADSWITCH", "COUPLING", "CIRC_BREAKER_2" };

        public string Datei, Netz, Planname = "";
        public DateTime? Planstand;
        public List<string> Hinweise = new List<string>();
        public List<Objekt> Alle = new List<Objekt>();       // ohne Klemmen (EDGE), in Dateireihenfolge
        public List<Objekt> Echt = new List<Objekt>();       // ohne Rechenparameter wie WGS84
        public Dictionary<int, Objekt> Nach = new Dictionary<int, Objekt>();

        private byte[] d;
        private List<Text> cs;
        private int[] starts;
        private readonly Dictionary<int, int[]> klemmen = new Dictionary<int, int[]>();
        private readonly Dictionary<int, List<int[]>> kanten = new Dictionary<int, List<int[]>>();
        private readonly Dictionary<int, int> abgaenge = new Dictionary<int, int>();
        private readonly List<int> abgangReihe = new List<int>();
        private readonly Dictionary<int, Text> abgangText = new Dictionary<int, Text>();
        private readonly HashSet<string> uwPraefix = new HashSet<string>();
        private readonly Dictionary<int, HashSet<int>> abgangKnoten = new Dictionary<int, HashSet<int>>();
        private readonly Dictionary<int, int> komp = new Dictionary<int, int>();
        private readonly Dictionary<int, string> zoneName = new Dictionary<int, string>();
        private readonly Dictionary<int, List<Text[]>> abschnitte = new Dictionary<int, List<Text[]>>();
        private List<Aenderung> aenderungen = new List<Aenderung>();

        private int I32(int p) { return BitConverter.ToInt32(d, p); }

        public static Vollexport Lies(string datei)
        {
            var v = new Vollexport();
            v.Datei = Path.GetFullPath(datei);
            v.Netz = Path.GetFileNameWithoutExtension(v.Datei).ToUpperInvariant();
            v.d = new Cfb(Leser.LiesDatei(v.Datei)).LiesStrom("Root/Data");
            v.cs = Leser.Texte(v.d);
            v.starts = new int[v.cs.Count];
            for (int i = 0; i < v.cs.Count; i++) v.starts[i] = v.cs[i].Start;
            v.Objekte();
            v.Anschluesse();
            v.AnLeitungsende();
            v.Filtern();
            v.Netzberechnung();
            v.Zonen();
            v.Abschnitte();
            Projekt p = Leser.Lies(v.Datei);
            v.aenderungen = p.Aenderungen;
            v.Planname = p.Planname;
            v.Planstand = Leser.LetzteAenderung(p);
            return v;
        }

        public int Anzahl(string typ)
        {
            int n = 0;
            foreach (Objekt o in Echt) if (o.Typ == typ) n++;
            return n;
        }

        // -------------------------------------------------------------- Objekte
        private void Objekte()
        {
            var nachEnde = new Dictionary<int, int>();
            for (int i = 0; i < cs.Count; i++) nachEnde[cs[i].Ende] = i;
            var tri = new List<int[]>();               // NEPID-Lage, NEPID, Namensindex, Typindex
            for (int i = 0; i < cs.Count; i++)
            {
                if (!TypMuster.IsMatch(cs[i].Wert)) continue;
                int j;
                if (!nachEnde.TryGetValue(cs[i].Start, out j)) continue;
                if (cs[j].Wert.Length == 0 || cs[j].Start < 4) continue;
                tri.Add(new[] { cs[j].Start - 4, I32(cs[j].Start - 4), j, i });
            }
            tri.Sort((x, y) => x[0].CompareTo(y[0]));
            for (int x = 0; x < tri.Count; x++)
            {
                string typ = cs[tri[x][3]].Wert;
                if (typ == "EDGE") continue;
                var o = new Objekt { Nepid = tri[x][1], NameIdx = tri[x][2], TypIdx = tri[x][3], Typ = typ, Name = cs[tri[x][2]].Wert };
                o.A = x > 0 ? cs[tri[x - 1][3]].Ende : 0;
                o.E = cs[tri[x][3]].Ende;
                int k0 = Untergrenze(o.A), k1 = Untergrenze(o.E);
                o.Texte = new int[Math.Max(0, k1 - k0)];
                for (int k = k0; k < k1; k++) o.Texte[k - k0] = k;
                o.NamePos = Array.IndexOf(o.Texte, o.NameIdx);
                Nach[o.Nepid] = o;                     // wie im Prototyp: das letzte Vorkommen gilt
            }
            foreach (Objekt o in Nach.Values) Alle.Add(o);
            Alle.Sort((a, b) => a.A.CompareTo(b.A));
        }

        private int Untergrenze(int pos)
        {
            int lo = 0, hi = starts.Length;
            while (lo < hi) { int m = (lo + hi) / 2; if (starts[m] < pos) lo = m + 1; else hi = m; }
            return lo;
        }

        // -------------------------------------------------------------- Regeln lesen
        private double? Zahl(Objekt o, string art, int k, int abstand, char fmt)
        {
            bool vorn = art == "v" || art == "ve";
            int kk = vorn ? k : k + o.NamePos;
            if (kk >= o.Texte.Length || kk < -1 || (!vorn && kk < 0)) return null;
            int basis;
            if (art == "v" || art == "h") basis = kk >= 0 ? cs[o.Texte[kk]].Ende : o.A;
            else
            {
                if (kk + 1 >= o.Texte.Length) return null;
                basis = cs[o.Texte[kk + 1]].Start;
            }
            return LiesZahl(basis + abstand, fmt);
        }

        private double? LiesZahl(int q, char fmt)
        {
            int n = fmt == 'd' ? 8 : fmt == 'h' ? 2 : fmt == 'b' ? 1 : 4;
            if (q < 0 || q + n > d.Length) return null;
            switch (fmt)
            {
                case 'd': return BitConverter.ToDouble(d, q);
                case 'f': return BitConverter.ToSingle(d, q);
                case 'i': return BitConverter.ToInt32(d, q);
                case 'h': return BitConverter.ToInt16(d, q);
                default: return d[q];
            }
        }

        private string TextVor(Objekt o, int k)
        {
            int kk = k + o.NamePos;
            return kk >= 0 && kk < o.Texte.Length ? cs[o.Texte[kk]].Wert : null;
        }

        private string TextVorn(Objekt o, int k)
        {
            return k >= 0 && k < o.Texte.Length ? cs[o.Texte[k]].Wert : null;
        }

        private object Wert(Objekt o, Regel r)
        {
            switch (r.Art)
            {
                case "null": return 0;
                case "leer": return null;
                case "tv": return TextVorn(o, r.K);
                case "th": return TextVor(o, r.K);
                default:
                    double? v = Zahl(o, r.Art, r.K, r.Abstand, r.Format);
                    return v.HasValue ? (object)v.Value : null;
            }
        }

        // -------------------------------------------------------------- Topologie
        private void Anschluesse()
        {
            var stat = new HashSet<int>();
            var andere = new HashSet<int>();
            foreach (Objekt o in Alle) (o.Typ == "BUSBAR-NODE" ? stat : andere).Add(o.Nepid);
            var unbekannt = new List<long[]>();
            for (int s = 0; s < 4; s++)
            {
                int n = (d.Length - s) / 4;
                for (int k = 1; k < n - 17; k++)
                {
                    int off = s + 4 * k;
                    if (I32(off - 4) != 2) continue;
                    int z = I32(off + 12);
                    if (z != 0 && z != 1) continue;
                    int a = I32(off), b = I32(off + 4), el, st;
                    if (andere.Contains(a) && stat.Contains(b)) { el = a; st = b; }
                    else if (stat.Contains(a) && andere.Contains(b)) { el = b; st = a; }
                    else continue;
                    int kl = I32(off + 8);
                    Objekt e = Nach[el];
                    if (kl < 0 || kl > 3)
                    {
                        // Im Umspannwerk steht im Klemmenfeld mitunter etwas anderes; bei
                        // Schaltgeraeten gehoert der Satz dann an die noch freie Klemme
                        if (Schaltgeraet.Contains(e.Typ)) unbekannt.Add(new long[] { kl, el, st, z });
                        continue;
                    }
                    if (e.Knoten[kl] == 0) { e.Knoten[kl] = st; e.Zustand[kl] = z; }
                    int kid = I32(off + 64);
                    if (!klemmen.ContainsKey(kid)) klemmen[kid] = new[] { el, kl };
                }
            }
            unbekannt.Sort((x, y) =>
            {
                for (int i = 0; i < 4; i++) { int c = x[i].CompareTo(y[i]); if (c != 0) return c; }
                return 0;
            });
            foreach (long[] u in unbekannt)
            {
                Objekt e = Nach[(int)u[1]];
                int st = (int)u[2];
                if (Array.IndexOf(e.Knoten, st) >= 0) continue;
                for (int i = 0; i < 2; i++)
                    if (e.Knoten[i] == 0) { e.Knoten[i] = st; e.Zustand[i] = (int)u[3]; break; }
            }
        }

        private Dictionary<string, int> StationNachName()
        {
            var r = new Dictionary<string, int>();
            foreach (Objekt o in Alle) if (o.Typ == "BUSBAR-NODE") r[o.Name.Trim()] = o.Nepid;
            return r;
        }

        private static bool Frei(Objekt o)
        {
            return o.Knoten[0] == 0 && o.Knoten[1] == 0 && o.Knoten[2] == 0 && o.Knoten[3] == 0;
        }

        // Schalter und Schutz an einem Leitungs-, Lastschalter- oder Einspeiseende: im Objekt
        // vor der NEPID steht die Klemmennummer dieses Endes.
        private void AnLeitungsende()
        {
            var stat = new HashSet<int>();
            foreach (Objekt o in Alle) if (o.Typ == "BUSBAR-NODE") stat.Add(o.Nepid);
            Dictionary<string, int> statNachName = StationNachName();
            foreach (Objekt o in Alle)
            {
                if (o.Typ == "VOL_RELAIS" && Frei(o))
                {
                    double? b = Zahl(o, "he", -7, -12, 'i');       // Spannungsrelais nennen ihre Sammelschiene
                    if (b.HasValue && stat.Contains((int)b.Value)) { o.Knoten[0] = (int)b.Value; o.Zustand[0] = 1; }
                    continue;
                }
                if ((o.Typ != "DISCSWITCH_2" && o.Typ != "CIRC_BREAKER_2" && o.Typ != "OVERCUR_RELAIS") || !Frei(o)) continue;
                int np = cs[o.NameIdx].Start - 4, kid = 0;
                bool gefunden = false;
                for (int q = np - 4; q >= o.A && q >= 0; q--)
                {
                    int v = I32(q);
                    if (klemmen.ContainsKey(v)) { kid = v; gefunden = true; break; }
                }
                if (gefunden)
                {
                    int[] ende = klemmen[kid];
                    Objekt el = Nach[ende[0]];
                    o.Anker = el.Nepid;
                    o.Knoten[1] = el.Knoten[ende[1]];
                    o.Zustand[0] = o.Zustand[1] = el.Zustand[ende[1]];
                    if (o.Typ == "OVERCUR_RELAIS") o.Zustand[0] = o.Zustand[1] = 1;   // Schutz ist kein Schalter
                }
                else if (o.Typ == "DISCSWITCH_2")
                {
                    string al = TextVor(o, -2);
                    int st;
                    if (al != null && statNachName.TryGetValue(al.Trim(), out st)) o.Knoten[1] = st;
                }
            }
        }

        private void Filtern()
        {
            var fremd = new Dictionary<string, int>();
            foreach (Objekt o in Alle)
            {
                bool bekannt = Tabelle.ContainsKey(o.Typ);
                if (!bekannt && Frei(o)) continue;
                Echt.Add(o);
                if (!bekannt) { int n; fremd.TryGetValue(o.Typ, out n); fremd[o.Typ] = n + 1; }
            }
            foreach (var kv in fremd)
                Hinweise.Add("Unbekannte Elementart " + kv.Key + " (" + kv.Value + " Stück), nur in TOPOLOGY übernommen.");
        }

        private void Kante(int a, int b, int zu)
        {
            List<int[]> l;
            if (!kanten.TryGetValue(a, out l)) { l = new List<int[]>(); kanten[a] = l; }
            l.Add(new[] { b, zu });
        }

        private bool Uw(int n)
        {
            string nm = Nach[n].Name;
            return nm.Contains("-") && uwPraefix.Contains(nm.Split('-')[0]);
        }

        private List<int> KnotenVon(Objekt o)
        {
            var l = new List<int>();
            if (o.Typ == "BUSBAR-NODE") l.Add(o.Nepid);
            else foreach (int k in o.Knoten) if (k != 0) l.Add(k);
            return l;
        }

        // Abgang und Teilnetz rechnet NEPLAN beim Export aus; hier dasselbe
        private void Netzberechnung()
        {
            foreach (Objekt o in Alle)
            {
                if (o.Typ == "BUSBAR-NODE" || o.Knoten[0] == 0 || o.Knoten[1] == 0 || o.Anker != 0) continue;
                int zu = o.Zustand[0] == 1 && o.Zustand[1] == 1 ? 1 : 0;
                Kante(o.Knoten[0], o.Knoten[1], zu);
                Kante(o.Knoten[1], o.Knoten[0], zu);
            }
            // Abgaenge: Feldname, 16 Bytes dahinter die NEPID der abgehenden Leitung
            Dictionary<string, int> statNachName = StationNachName();
            foreach (Text x in cs)
            {
                int n;
                if (!statNachName.TryGetValue(x.Wert.Trim(), out n) || x.Ende + 20 > d.Length) continue;
                int el = I32(x.Ende + 16);
                Objekt l;
                if (Nach.TryGetValue(el, out l) && l.Typ == "LINE" && (l.Knoten[0] == n || l.Knoten[1] == n))
                {
                    if (!abgaenge.ContainsKey(n)) abgangReihe.Add(n);
                    abgaenge[n] = el;
                    abgangText[n] = x;
                }
            }
            foreach (int n in abgangReihe) uwPraefix.Add(Nach[n].Name.Split('-')[0]);
            foreach (int f in abgangReihe)
            {
                var gesehen = new HashSet<int> { f };
                var warte = new Queue<int>();
                warte.Enqueue(f);
                while (warte.Count > 0)
                {
                    List<int[]> nb;
                    if (!kanten.TryGetValue(warte.Dequeue(), out nb)) continue;
                    foreach (int[] m in nb)
                        if (m[1] == 1 && !gesehen.Contains(m[0]) && !Uw(m[0])) { gesehen.Add(m[0]); warte.Enqueue(m[0]); }
                }
                foreach (int x in gesehen)
                {
                    HashSet<int> h;
                    if (!abgangKnoten.TryGetValue(x, out h)) { h = new HashSet<int>(); abgangKnoten[x] = h; }
                    h.Add(f);
                }
            }
            // Teilnetze: Zusammenhangskomponenten, nummeriert in Datei-Reihenfolge der Elemente
            int nr = 0;
            foreach (Objekt o in Echt)
            {
                foreach (int n in KnotenVon(o))
                {
                    if (komp.ContainsKey(n)) continue;
                    nr++;
                    komp[n] = nr;
                    var warte = new Queue<int>();
                    warte.Enqueue(n);
                    while (warte.Count > 0)
                    {
                        List<int[]> nb;
                        if (!kanten.TryGetValue(warte.Dequeue(), out nb)) continue;
                        foreach (int[] m in nb)
                            if (m[1] == 1 && !komp.ContainsKey(m[0])) { komp[m[0]] = nr; warte.Enqueue(m[0]); }
                    }
                }
            }
        }

        private string Abgang(Objekt o)
        {
            if (o.Anker != 0) return o.Zustand[0] == 0 ? "" : Abgang(Nach[o.Anker]);
            var alle = new List<int>();
            var zu = new List<int>();
            if (o.Typ == "BUSBAR-NODE") { alle.Add(o.Nepid); zu.Add(o.Nepid); }
            else
                for (int i = 0; i < 4; i++)
                    if (o.Knoten[i] != 0) { alle.Add(o.Knoten[i]); if (o.Zustand[i] == 1) zu.Add(o.Knoten[i]); }
            if (alle.Count > 0)
            {
                bool imUw = true;
                foreach (int k in alle) if (!Uw(k)) { imUw = false; break; }
                if (imUw) return "";
            }
            var s = new HashSet<int>();
            foreach (int k in zu) { HashSet<int> h; if (abgangKnoten.TryGetValue(k, out h)) s.UnionWith(h); }
            if (s.Count != 1) return "";
            foreach (int f in s) return Nach[f].Name;
            return "";
        }

        private int Teilnetz(Objekt o)
        {
            int v;
            if (o.Typ == "BUSBAR-NODE") return komp.TryGetValue(o.Nepid, out v) ? v : -1;
            if (o.Anker != 0) return Teilnetz(Nach[o.Anker]);
            for (int i = 0; i < 4; i++)
                if (o.Knoten[i] != 0 && o.Zustand[i] == 1 && komp.TryGetValue(o.Knoten[i], out v)) return v;
            return -1;
        }

        // Zonen: Lasten und Erzeuger tragen die ID ihrer Zone, der Zonenname steht hinter dieser ID
        private static readonly Dictionary<string, Regel> ZoneRegel = new Dictionary<string, Regel> {
            { "LOAD", new Regel("", "", "h", -6, 0, 'i') }, { "AC_GENERIC_COMP", new Regel("", "", "v", 8, 0, 'i') } };

        private void Zonen()
        {
            var ids = new HashSet<int>();
            foreach (Objekt o in Alle)
            {
                Regel r;
                if (!ZoneRegel.TryGetValue(o.Typ, out r)) continue;
                double? v = Zahl(o, r.Art, r.K, r.Abstand, r.Format);
                if (v.HasValue) ids.Add((int)v.Value);
            }
            foreach (Text x in cs)
            {
                if (x.Start < 4 || x.Wert.Length == 0) continue;
                int id = I32(x.Start - 4);
                if (ids.Contains(id) && !zoneName.ContainsKey(id)) zoneName[id] = x.Wert;
            }
        }

        private string Zone(Objekt o)
        {
            Regel r;
            if (!ZoneRegel.TryGetValue(o.Typ, out r)) return "Zone 1";
            double? v = Zahl(o, r.Art, r.K, r.Abstand, r.Format);
            string n;
            return v.HasValue && zoneName.TryGetValue((int)v.Value, out n) ? n : "";
        }

        // Kabelabschnitte: Text Kabeltyp | 120 Bytes | Text | 16 Bytes | Text Materialnummer
        private void Abschnitte()
        {
            foreach (Objekt o in Alle)
            {
                if (o.Typ != "LINE") continue;
                var l = new List<Text[]>();
                int[] t = o.Texte;
                for (int p = 0; p + 2 < t.Length; p++)
                    if (cs[t[p + 1]].Start - cs[t[p]].Ende == 120 && cs[t[p + 2]].Start - cs[t[p + 1]].Ende == 16)
                        l.Add(new[] { cs[t[p]], cs[t[p + 1]], cs[t[p + 2]] });
                abschnitte[o.Nepid] = l;
            }
        }

        // -------------------------------------------------------------- Schreiben
        // Das Ziel muss eine frische Kopie der vollen Vorlage sein.
        public Dictionary<string, int> Schreibe(string mdb)
        {
            try { return SchreibeIntern(mdb); }
            finally { Mdb.Freigeben(); }
        }

        private Dictionary<string, object> Grund(Objekt o)
        {
            return new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase) { { "NETNAME", Netz }, { "NEPID", o.Nepid }, { "NAME", o.Name } };
        }

        private Dictionary<string, object> Zeile(params object[] paare)
        {
            var z = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase) { { "NETNAME", Netz } };
            for (int i = 0; i + 1 < paare.Length; i += 2) z[(string)paare[i]] = paare[i + 1];
            return z;
        }

        private Dictionary<string, int> SchreibeIntern(string mdb)
        {
            var anzahl = new Dictionary<string, int>();
            using (OleDbConnection c = Mdb.Oeffne(mdb, true))
            {
                if (!Mdb.HatTabelle(c, "A_AENDERUNGEN"))
                    new OleDbCommand("CREATE TABLE [A_AENDERUNGEN] ([NETNAME] TEXT(250), [LFD] INTEGER, [REV] TEXT(10), [NUMMER] TEXT(50), " +
                                     "[DATUM] TEXT(20), [KUERZEL] TEXT(50), [PLANNAME] TEXT(250))", c).ExecuteNonQuery();
                var schema = Schema(c);
                var regelnJe = new Dictionary<string, List<Regel>>();
                foreach (Regel r in Regeln.Spalten)
                {
                    List<Regel> l;
                    if (!regelnJe.TryGetValue(r.Tabelle, out l)) { l = new List<Regel>(); regelnJe[r.Tabelle] = l; }
                    l.Add(r);
                }
                using (OleDbTransaction tx = c.BeginTransaction())
                {
                    var w = new Schreiber(c, tx, schema, anzahl);
                    foreach (Objekt o in Echt)
                    {
                        string tab;
                        if (Tabelle.TryGetValue(o.Typ, out tab) && regelnJe.ContainsKey(tab))
                        {
                            var wv = Grund(o);
                            foreach (Regel r in regelnJe[tab]) wv[r.Spalte] = Wert(o, r);
                            if (o.Typ == "DISCSWITCH_2") wv["ALIASNAME"] = TextVor(o, -2);
                            if (o.Typ == "BUSBAR-NODE")
                                for (int kk = o.NamePos - 1; kk > Math.Max(-1, o.NamePos - 15); kk--)
                                    if (cs[o.Texte[kk]].Wert == "panelname") { wv["panelname"] = cs[o.Texte[kk + 1]].Wert; break; }
                            w.Zeile(tab, wv);
                        }
                        var tw = Grund(o);
                        tw["TYPE"] = o.Typ;
                        tw["NETAREA"] = "Area 1";
                        tw["NETZONE"] = Zone(o);
                        tw["PARTNET"] = Teilnetz(o);
                        tw["FEEDER"] = Abgang(o);
                        tw["DEL"] = 0;
                        tw["CHANGE_NEW"] = 0;
                        for (int i = 0; i < 4; i++)
                        {
                            tw["NODE" + (i + 1)] = o.Knoten[i] != 0 ? (object)Nach[o.Knoten[i]].Name : null;
                            tw["SW" + (i + 1)] = o.Knoten[i] != 0 ? (object)o.Zustand[i] : null;
                        }
                        if (o.Anker != 0)
                        {
                            tw["NODE1"] = Nach[o.Anker].Name;
                            tw["SW1"] = o.Zustand[0];
                            tw["NODE3"] = Nach[o.Anker].Typ;
                        }
                        else if (o.Typ == "VOL_RELAIS" && o.Knoten[0] != 0) tw["NODE3"] = "BUSBAR-NODE";
                        w.Zeile("TOPOLOGY", tw);
                        if (o.Typ == "LINE") Abschnittszeilen(o, w);
                    }
                    Farben(w);
                    Faktoren(w);
                    Benutzerdaten(w);
                    w.Zeile("INFOTABLE", Zeile("VERSION", Version, "LIB_PROJECT", 1, "PROJECT", Datei, "VARIANT", "Rootnet",
                        "NAMES_IDS", 0, "DELETE_UPDATE", 0, "ALL_CHANGED", 0, "EXPORTNODES", 1, "EXPORTGRAPHIC", 0,
                        "EXPORT_TIME", DateTime.Now, "WORLD_LOG", 0, "USE42SYMBOLS", 0, "METER_BAR_NODE", 0, "METER_BAR_ELEM", 0,
                        "EXPORT_ONLY_FEEDED", 0, "USE_DEFAULT_COMPONENT_ORDER", 0, "IMPORT_AUXILIARY_GRAPHIC", 1,
                        "DELETE_EMPTY_LAYERS", 0, "IMPORT_EXPORT_LABELS", 1, "PROTECTIONAUTOASSIGN", 1));
                    for (int i = 0; i < aenderungen.Count; i++)
                        w.Zeile("A_AENDERUNGEN", Zeile("LFD", i + 1, "REV", aenderungen[i].Rev, "NUMMER", aenderungen[i].Nummer,
                            "DATUM", aenderungen[i].Datum, "KUERZEL", aenderungen[i].Kuerzel, "PLANNAME", Planname));
                    tx.Commit();
                }
            }
            return anzahl;
        }

        private void Abschnittszeilen(Objekt o, Schreiber w)
        {
            List<Text[]> l;
            if (!abschnitte.TryGetValue(o.Nepid, out l)) return;
            for (int i = 0; i < l.Count; i++)
            {
                var sw = Grund(o);
                sw["NUMLINE"] = i + 1;
                foreach (Regel r in Regeln.Abschnitte)
                {
                    if (r.Art == "null") sw[r.Spalte] = 0;
                    else if (r.Art == "leer") sw[r.Spalte] = null;
                    else if (r.Art == "abt") sw[r.Spalte] = l[i][r.K].Wert;
                    else
                    {
                        double? v = LiesZahl(l[i][r.K].Ende + r.Abstand, r.Format);
                        sw[r.Spalte] = v.HasValue ? (object)v.Value : null;
                    }
                }
                w.Zeile("LINESECTIONS", sw);
            }
        }

        private static string Farbe(int v) { return v == 0 ? "0" : "0X" + ((uint)v).ToString("X"); }

        private int Suche(byte[] muster)
        {
            for (int i = 0; i + muster.Length <= d.Length; i++)
            {
                int k = 0;
                while (k < muster.Length && d[i + k] == muster[k]) k++;
                if (k == muster.Length) return i;
            }
            return -1;
        }

        private void Farben(Schreiber w)
        {
            // Spannungsebenen: hinter der Klassenkennung je Objekt 3 | Farbe | Spannung | GUID-Text
            int kennung = Suche(Encoding.ASCII.GetBytes("CNPVoltageLevel"));
            if (kennung >= 0)
            {
                int p = kennung + 15 + 4;
                while (p + 16 <= d.Length && I32(p - 4) == 3 && d[p + 12] == 0xFF && d[p + 13] == 0xFE && d[p + 14] == 0xFF && d[p + 15] != 0xFF)
                {
                    w.Zeile("NETWORKCOLORS", Zeile("COLORTYPE", "VOLTAGE_LEVEL", "COLOR", Farbe(I32(p)), "UN", BitConverter.ToDouble(d, p + 4)));
                    p = p + 16 + 2 * d[p + 15] + 10;
                }
            }
            // Zonen: 9 | ID | Name | Farbe | aktiv | 1 - Bereich
            foreach (Text x in cs)
            {
                if (x.Start < 8 || x.Ende + 12 > d.Length || I32(x.Start - 8) != 9) continue;
                int f = I32(x.Ende), a = I32(x.Ende + 4), b = I32(x.Ende + 8);
                if (f >= 0 && f <= 0xFFFFFF && (a == 0 || a == 1) && (b == 0 || b == 1))
                    w.Zeile("NETWORKCOLORS", Zeile("COLORTYPE", "ZONE", "COLOR", Farbe(f), "DESCRIPTION", x.Wert, "ISACTIVE", a, "ISAREA", 1 - b));
            }
            // Teilnetze: Nummerntext | Farbe | Nummer | 0 | 0/1 | 0, naechster Text genau 36 Bytes weiter
            var anfaenge = new HashSet<int>(starts);
            foreach (Text x in cs)
            {
                string t = x.Wert.Trim();
                int nr;
                if (t.Length == 0 || !int.TryParse(t, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out nr)) continue;
                if (x.Ende + 24 > d.Length || !anfaenge.Contains(x.Ende + 36)) continue;
                int f = I32(x.Ende), n2 = I32(x.Ende + 4), z0 = I32(x.Ende + 8), fl = I32(x.Ende + 12), z2 = I32(x.Ende + 16);
                if (n2 == nr && f >= 0 && f <= 0xFFFFFF && z0 == 0 && z2 == 0 && (fl == 0 || fl == 1))
                    w.Zeile("NETWORKCOLORS", Zeile("COLORTYPE", "PARTIAL_NET", "COLOR", Farbe(f), "PARTNETNR", nr));
            }
            // Zeichenebenen: 1 | x | 1 | 0 | Name | leerer Text | 0 | Farbe
            foreach (Text x in cs)
            {
                if (x.Start < 16 || x.Ende + 12 > d.Length) continue;
                if (I32(x.Start - 16) == 1 && I32(x.Start - 8) == 1 && I32(x.Start - 4) == 0 && I32(x.Ende) == 16776959 && I32(x.Ende + 4) == 0)
                    w.Zeile("NETWORKCOLORS", Zeile("COLORTYPE", "GRAPHIC_LAYER", "COLOR", Farbe(I32(x.Ende + 8)), "DESCRIPTION", x.Wert, "ISACTIVE", 0));
            }
            // Abgaenge: Farbe direkt hinter dem Feldnamen
            foreach (int n in abgangReihe)
            {
                Objekt l = Nach[abgaenge[n]];
                w.Zeile("NETWORKCOLORS", Zeile("COLORTYPE", "FEEDER", "COLOR", Farbe(I32(abgangText[n].Ende)), "DESCRIPTION", Nach[n].Name,
                    "FEEDERNODEID", n, "FEEDERELEMID", l.Nepid, "FEEDERNODENAME", Nach[n].Name, "FEEDERELEMNAME", l.Name, "ELEMENTTYPE", l.Typ));
            }
        }

        private void Faktoren(Schreiber w)
        {
            // P und Q stehen VOR dem eigenen Block: 68 bzw. 60 Bytes vor dem Text "SCALING_TYPES"
            var namen = new HashSet<string>();
            for (int i = 0; i + 1 < cs.Count; i++)
            {
                if (cs[i].Wert != "SCALING_TYPES") continue;
                namen.Add(cs[i + 1].Wert);
                int p = cs[i].Start;
                w.Zeile("SCALINGFACTORS", Zeile("NAME", cs[i + 1].Wert,
                    "FACTORP", p >= 68 ? (object)BitConverter.ToDouble(d, p - 68) : null,
                    "FACTORQ", p >= 60 ? (object)BitConverter.ToDouble(d, p - 60) : null));
            }
            // Zuordnung: die Faktornamen im Last- bzw. Erzeugerobjekt
            foreach (Objekt o in Echt)
            {
                if (o.Typ != "LOAD" && o.Typ != "AC_GENERIC_COMP") continue;
                foreach (int k in o.Texte)
                    if (namen.Contains(cs[k].Wert))
                        w.Zeile("SCALINGFACTORSASSIGN", Zeile("NEPID", o.Nepid, "NAME", o.Name, "SCALINGFACTOR", cs[k].Wert, "PORTION", 100));
            }
        }

        // Benutzerdaten: hinter dem Typnamen (22 Bytes) Paare Beschriftung | 8 | Wert, Paare mit 18 Bytes Abstand
        private List<string> Felder(Objekt o)
        {
            var l = new List<string>();
            int[] t = o.Texte;
            for (int p = 0; p + 1 < t.Length; p++)
            {
                if (cs[t[p]].Wert != o.Typ || cs[t[p + 1]].Start - cs[t[p]].Ende != 22) continue;
                int q = p + 1;
                while (q + 1 < t.Length && cs[t[q + 1]].Start - cs[t[q]].Ende == 8)
                {
                    l.Add(cs[t[q]].Wert);
                    if (q + 2 < t.Length && cs[t[q + 2]].Start - cs[t[q + 1]].Ende == 18) q += 2;
                    else break;
                }
                break;
            }
            return l;
        }

        private void Benutzerdaten(Schreiber w)
        {
            var typen = new List<string>();
            foreach (Objekt o in Echt) if (!typen.Contains(o.Typ)) typen.Add(o.Typ);
            typen.Sort(StringComparer.Ordinal);
            foreach (string typ in typen)
            {
                var formen = new Dictionary<string, int>();
                var reihe = new List<string>();
                var inhalt = new Dictionary<string, List<string>>();
                foreach (Objekt o in Echt)
                {
                    if (o.Typ != typ) continue;
                    List<string> l = Felder(o);
                    if (l.Count == 0) continue;
                    string key = string.Join("\u0001", l.ToArray());
                    int n;
                    if (!formen.TryGetValue(key, out n)) { reihe.Add(key); inhalt[key] = l; }
                    formen[key] = n + 1;
                }
                string beste = null;
                foreach (string key in reihe) if (beste == null || formen[key] > formen[beste]) beste = key;   // haeufigste Form
                if (beste == null) continue;
                List<string> labels = inhalt[beste];
                for (int i = 0; i < labels.Count; i++)
                    w.Zeile("USERDATADEF", Zeile("ELEMTYPE", typ, "VARID", labels.Count - i, "VARTYPE", 40, "NAME", labels[i], "DISPVALUE", 0, "DISPDESCR", 0));
            }
        }

        private static Dictionary<string, Dictionary<string, int>> Schema(OleDbConnection c)
        {
            var s = new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);
            foreach (DataRow r in c.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, null).Rows)
            {
                string t = (string)r["TABLE_NAME"];
                if (t.StartsWith("MSys", StringComparison.OrdinalIgnoreCase)) continue;
                Dictionary<string, int> l;
                if (!s.TryGetValue(t, out l)) { l = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase); s[t] = l; }
                l[(string)r["COLUMN_NAME"]] = Convert.ToInt32(r["DATA_TYPE"]);
            }
            return s;
        }

        private static double Zahlwert(object v)
        {
            double x;
            if (v is double) x = (double)v;
            else if (v is string) { if (!double.TryParse((string)v, NumberStyles.Float, CultureInfo.InvariantCulture, out x)) x = 0; }
            else x = Convert.ToDouble(v, CultureInfo.InvariantCulture);
            return double.IsNaN(x) || double.IsInfinity(x) ? 0 : x;
        }

        // Wert in den Spaltentyp der Vorlage bringen (OLE-DB-Typcodes)
        private static object Wandle(object v, int typ)
        {
            if (v == null) return typ == 11 ? (object)false : DBNull.Value;
            switch (typ)
            {
                case 130: case 202: case 203:
                    return v is double ? ((double)v).ToString("R", CultureInfo.InvariantCulture) : Convert.ToString(v, CultureInfo.InvariantCulture);
                case 11: return Zahlwert(v) != 0;
                case 7: return v is DateTime ? v : (object)DBNull.Value;
                case 2: return (short)Math.Round(Zahlwert(v));
                case 3: return (int)Math.Round(Zahlwert(v));
                case 17: return (byte)Math.Round(Zahlwert(v));
                case 4: return (float)Zahlwert(v);
                default: return Zahlwert(v);
            }
        }

        private static OleDbParameter Param(object v, int typ)
        {
            var p = new OleDbParameter();
            switch (typ)
            {
                case 130: case 202: p.OleDbType = OleDbType.VarWChar; break;
                case 203: p.OleDbType = OleDbType.LongVarWChar; break;
                case 11: p.OleDbType = OleDbType.Boolean; break;
                case 7: p.OleDbType = OleDbType.Date; break;
                case 2: p.OleDbType = OleDbType.SmallInt; break;
                case 3: p.OleDbType = OleDbType.Integer; break;
                case 17: p.OleDbType = OleDbType.UnsignedTinyInt; break;
                case 4: p.OleDbType = OleDbType.Single; break;
                default: p.OleDbType = OleDbType.Double; break;
            }
            p.Value = v;
            return p;
        }

        private class Schreiber
        {
            private readonly OleDbConnection c;
            private readonly OleDbTransaction tx;
            private readonly Dictionary<string, Dictionary<string, int>> schema;
            private readonly Dictionary<string, int> anzahl;

            public Schreiber(OleDbConnection c, OleDbTransaction tx, Dictionary<string, Dictionary<string, int>> schema, Dictionary<string, int> anzahl)
            {
                this.c = c; this.tx = tx; this.schema = schema; this.anzahl = anzahl;
            }

            // Ein vorbereiteter Befehl je Tabelle und Spaltensatz, fuer jede Zeile nur die Werte neu
            private readonly Dictionary<string, OleDbCommand> befehle = new Dictionary<string, OleDbCommand>();

            public void Zeile(string tabelle, Dictionary<string, object> werte)
            {
                Dictionary<string, int> spalten;
                if (!schema.TryGetValue(tabelle, out spalten)) return;
                var namen = new List<string>();
                var typen = new List<int>();
                var inhalt = new List<object>();
                foreach (var kv in werte)
                {
                    int typ;
                    if (!spalten.TryGetValue(kv.Key, out typ)) continue;
                    namen.Add(kv.Key); typen.Add(typ); inhalt.Add(kv.Value);
                }
                string schl = tabelle + "|" + string.Join("|", namen.ToArray());
                OleDbCommand cmd;
                if (!befehle.TryGetValue(schl, out cmd))
                {
                    var platz = new string[namen.Count];
                    for (int i = 0; i < platz.Length; i++) platz[i] = "?";
                    cmd = new OleDbCommand("INSERT INTO [" + tabelle + "] ([" + string.Join("], [", namen.ToArray()) + "]) VALUES (" +
                                           string.Join(", ", platz) + ")", c, tx);
                    for (int i = 0; i < namen.Count; i++) cmd.Parameters.Add(Param(null, typen[i]));
                    befehle[schl] = cmd;
                }
                for (int i = 0; i < namen.Count; i++) cmd.Parameters[i].Value = Wandle(inhalt[i], typen[i]);
                cmd.ExecuteNonQuery();
                int n;
                anzahl.TryGetValue(tabelle, out n);
                anzahl[tabelle] = n + 1;
            }
        }
    }

    // ------------------------------------------------------------------ //
    // Zellvergleich: NEPLAN-Export gegen Export ohne NEPLAN, ueber alle Tabellen
    // ------------------------------------------------------------------ //
    public class Zellvergleich
    {
        public List<string> Zeilen = new List<string>();
        public int Zellen, Gleich, Fehlt, Falsch, ZeilenZuviel;

        private static readonly Dictionary<string, string[]> Schluessel = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase) {
            { "LINESECTIONS", new[] { "NEPID", "NUMLINE" } }, { "SCALINGFACTORS", new[] { "NAME" } },
            { "SCALINGFACTORSASSIGN", new[] { "NEPID", "SCALINGFACTOR" } }, { "USERDATADEF", new[] { "ELEMTYPE", "NAME" } },
            { "INFOTABLE", new[] { "NETNAME" } },
            { "NETWORKCOLORS", new[] { "COLORTYPE", "DESCRIPTION", "UN", "PARTNETNR", "FEEDERNODENAME" } },
            { "A_KnotenTextOffset", new[] { "SYMBOL_NAME", "EQART" } }, { "LASTMuster", new[] { "LFDNR", "coord_nr" } } };
        private static readonly HashSet<string> Unwichtig = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
            "INFOTABLE.EXPORT_TIME", "INFOTABLE.VERSION", "INFOTABLE.PROJECT" };

        public static Zellvergleich Vergleiche(string neplanMdb, string leserMdb, string netz)
        {
            try { return VergleicheIntern(neplanMdb, leserMdb, netz); }
            finally { Mdb.Freigeben(); }
        }

        private static object Norm(object v)
        {
            if (v == null || v == DBNull.Value) return null;
            if (v is bool) return (bool)v ? 1.0 : 0.0;
            if (v is string)
            {
                string s = ((string)v).Replace("\r\n", "\n").Replace("\r", "\n").Trim();
                if (s.Length == 0) return null;
                double x;
                if (double.TryParse(s.Replace(",", "."), NumberStyles.Float, CultureInfo.InvariantCulture, out x)) return x;
                return s;
            }
            if (v is DateTime) return ((DateTime)v).ToString("s");
            try { return Convert.ToDouble(v, CultureInfo.InvariantCulture); } catch { return Convert.ToString(v); }
        }

        private static bool Passt(object a, object b)
        {
            if (a == null || b == null) return a == null && b == null;
            if (a is double && b is double)
            {
                double x = (double)a, y = (double)b;
                return Math.Abs(x - y) <= Math.Max(1e-9, 1e-6 * Math.Max(Math.Abs(x), Math.Abs(y)));
            }
            return Equals(a, b);
        }

        private static Dictionary<string, List<DataRow>> Lies(OleDbConnection c, string netz, out Dictionary<string, List<string>> spalten)
        {
            var r = new Dictionary<string, List<DataRow>>(StringComparer.OrdinalIgnoreCase);
            spalten = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            foreach (DataRow t in c.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" }).Rows)
            {
                string name = (string)t["TABLE_NAME"];
                var dt = new DataTable();
                new OleDbDataAdapter("SELECT * FROM [" + name + "]", c).Fill(dt);
                var l = new List<DataRow>();
                bool mitNetz = dt.Columns.Contains("NETNAME");
                foreach (DataRow z in dt.Rows)
                    if (!mitNetz || string.Equals(Convert.ToString(z["NETNAME"]).Trim(), netz, StringComparison.OrdinalIgnoreCase)) l.Add(z);
                r[name] = l;
                var sp = new List<string>();
                foreach (DataColumn col in dt.Columns) sp.Add(col.ColumnName);
                spalten[name] = sp;
            }
            return r;
        }

        private static string Key(DataRow z, string[] key)
        {
            var sb = new StringBuilder();
            foreach (string k in key)
            {
                object v = z.Table.Columns.Contains(k) ? Norm(z[k]) : null;
                sb.Append(v == null ? "\u0001" : v is double ? ((double)v).ToString("R", CultureInfo.InvariantCulture) : v.ToString()).Append('\u0001');
            }
            return sb.ToString();
        }

        private static Zellvergleich VergleicheIntern(string neplanMdb, string leserMdb, string netz)
        {
            var v = new Zellvergleich();
            Dictionary<string, List<string>> rsp, psp;
            Dictionary<string, List<DataRow>> refd, prob;
            using (OleDbConnection c = Mdb.Oeffne(neplanMdb, false)) refd = Lies(c, netz, out rsp);
            using (OleDbConnection c = Mdb.Oeffne(leserMdb, false)) prob = Lies(c, netz, out psp);
            v.Zeilen.Add(string.Format("{0,-22}{1,12}{2,8}{3,8}{4,7}{5,7}   Spalten mit Problemen (fehlt/falsch)", "Tabelle", "Zeilen N/L", "Zellen", "gleich", "fehlt", "falsch"));
            var namen = new List<string>(refd.Keys);
            namen.Sort(StringComparer.OrdinalIgnoreCase);
            foreach (string t in namen)
            {
                List<DataRow> rz = refd[t];
                if (rz.Count == 0) continue;
                List<DataRow> pz;
                if (!prob.TryGetValue(t, out pz)) pz = new List<DataRow>();
                string[] key;
                if (!Schluessel.TryGetValue(t, out key)) key = rsp[t].Contains("NEPID") ? new[] { "NEPID" } : null;
                var paare = new List<DataRow[]>();
                if (key != null)
                {
                    var idx = new Dictionary<string, Queue<DataRow>>();
                    foreach (DataRow z in pz)
                    {
                        string k = Key(z, key);
                        Queue<DataRow> q;
                        if (!idx.TryGetValue(k, out q)) { q = new Queue<DataRow>(); idx[k] = q; }
                        q.Enqueue(z);
                    }
                    foreach (DataRow z in rz)
                    {
                        Queue<DataRow> q;
                        paare.Add(new[] { z, idx.TryGetValue(Key(z, key), out q) && q.Count > 0 ? q.Dequeue() : null });
                    }
                }
                else
                    for (int i = 0; i < rz.Count; i++) paare.Add(new[] { rz[i], i < pz.Count ? pz[i] : null });
                int zellen = 0, gleich = 0, fehlt = 0, falsch = 0;
                var probleme = new SortedDictionary<string, int[]>();
                foreach (DataRow[] p in paare)
                    foreach (string s in rsp[t])
                    {
                        if (Unwichtig.Contains(t + "." + s)) continue;
                        zellen++;
                        object a = Norm(p[0][s]);
                        object b = p[1] != null && p[1].Table.Columns.Contains(s) ? Norm(p[1][s]) : null;
                        if (Passt(a, b)) { gleich++; continue; }
                        int[] pr;
                        if (!probleme.TryGetValue(s, out pr)) { pr = new int[2]; probleme[s] = pr; }
                        if (b == null) { fehlt++; pr[0]++; } else { falsch++; pr[1]++; }
                    }
                int zuviel = Math.Max(0, pz.Count - rz.Count);
                v.ZeilenZuviel += zuviel;
                var text = new StringBuilder();
                foreach (var kv in probleme)
                {
                    if (text.Length > 90) { text.Append(" …"); break; }
                    text.Append(kv.Key).Append(' ').Append(kv.Value[0]).Append('/').Append(kv.Value[1]).Append(", ");
                }
                v.Zeilen.Add(string.Format("{0,-22}{1,6}/{2,-5}{3,8}{4,8}{5,7}{6,7}   {7}{8}", t, rz.Count, pz.Count, zellen, gleich, fehlt, falsch,
                    text.ToString().TrimEnd(' ', ','), zuviel > 0 ? "  (" + zuviel + " Zeilen zu viel)" : ""));
                v.Zellen += zellen; v.Gleich += gleich; v.Fehlt += fehlt; v.Falsch += falsch;
            }
            v.Zeilen.Add(string.Format("Gesamt: {0} Zellen, gleich {1} ({2:0.0} %), fehlt {3}, falsch {4}, Zeilen zu viel {5}",
                v.Zellen, v.Gleich, v.Zellen == 0 ? 0 : 100.0 * v.Gleich / v.Zellen, v.Fehlt, v.Falsch, v.ZeilenZuviel));
            return v;
        }
    }
}
