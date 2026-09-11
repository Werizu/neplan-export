// Schreibt ein gelesenes NEPLAN-Projekt in eine Access-Datenbank im Format des
// NEPLAN-Datenbankexports und vergleicht es mit einem echten NEPLAN-Export.
//
// Gefuellt wird nur, was der Leser verlaesslich liefert (am 11.09.2026 an Dingden
// gegen den NEPLAN-Export geprueft). Fehlende Tabellen und Spalten werden angelegt,
// wie NEPLAN es beim Export auch tut. Braucht die Jet-Schnittstelle, also einen
// 32-Bit-Prozess.

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Text;

namespace NeplanLeser
{
    public static class Mdb
    {
        public static OleDbConnection Oeffne(string mdb, bool schreiben)
        {
            Exception letzter = null;
            foreach (string prov in new[] { "Microsoft.Jet.OLEDB.4.0", "Microsoft.ACE.OLEDB.12.0", "Microsoft.ACE.OLEDB.16.0" })
            {
                try
                {
                    // OLE DB Services=-4: keine Wiederverwendung im Hintergrund. Sonst haelt der
                    // Prozess die Datenbank nach dem Schliessen offen (Sperrdatei .ldb bleibt),
                    // und niemand kann sie oeffnen, bis das Fenster zu ist.
                    var c = new OleDbConnection("Provider=" + prov + ";Data Source=" + mdb + ";OLE DB Services=-4;" + (schreiben ? "" : "Mode=Read;"));
                    c.Open();
                    return c;
                }
                catch (Exception ex) { letzter = ex; }
            }
            throw new InvalidOperationException("Die Datenbank lässt sich nicht öffnen: " + mdb +
                ". Das Werkzeug bitte über den Starter (.cmd) aufrufen. " + (letzter == null ? "" : letzter.Message));
        }

        // Jet haelt eine Datenbank nach dem Schliessen noch offen (Sperrdatei .ldb bleibt),
        // bis die COM-Objekte weggeraeumt sind. Das hier erzwingt es, damit andere sie
        // sofort oeffnen koennen, auch waehrend das Werkzeug noch laeuft.
        public static void Freigeben()
        {
            OleDbConnection.ReleaseObjectPool();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        public static bool HatTabelle(OleDbConnection c, string tabelle)
        {
            return c.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, tabelle, "TABLE" }).Rows.Count > 0;
        }

        public static HashSet<string> Spalten(OleDbConnection c, string tabelle)
        {
            var r = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (DataRow z in c.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, tabelle, null }).Rows)
                r.Add((string)z["COLUMN_NAME"]);
            return r;
        }
    }

    public static class MdbSchreiber
    {
        public static readonly Dictionary<string, string> Tabelle = new Dictionary<string, string> {
            { "BUSBAR-NODE", "BUSBAR" }, { "LINE", "LINE" }, { "LOAD", "LOAD" }, { "AC_GENERIC_COMP", "AC_GENERIC_COMP" },
            { "DISCSWITCH_2", "DISC_SWITCH" }, { "LOADSWITCH", "LOAD_SWITCH" }, { "CIRC_BREAKER_2", "CIRC_BREAK" },
            { "TRANSFORMER", "TRAFO2" }, { "FEEDER", "NETWFEEDER" }, { "COUPLING", "BB_COUPLER" },
            { "OVERCUR_RELAIS", "RELAY_OVERCUR" }, { "VOL_RELAIS", "RELAY_MINMAX_ON_NODE" } };

        // Nur Felder, die gegen den NEPLAN-Export geprueft sind
        private static readonly Dictionary<string, string[]> Felder = new Dictionary<string, string[]> {
            { "BUSBAR", new[] { "ALIASNAME", "DESCRIPTION", "HYPERLINK", "legacy", "panelname" } },
            { "LINE", new[] { "HYPERLINK", "legacy" } },
            { "LOAD", new[] { "ALIASNAME", "DESCRIPTION" } },
            { "AC_GENERIC_COMP", new[] { "ALIASNAME" } },
            { "DISC_SWITCH", new[] { "ALIASNAME", "DESCRIPTION" } },
            { "LOAD_SWITCH", new[] { "DESCRIPTION" } },
            { "CIRC_BREAK", new[] { "DESCRIPTION" } },
            { "BB_COUPLER", new[] { "DESCRIPTION" } },
            { "TRAFO2", new string[0] },
            { "NETWFEEDER", new[] { "ALIASNAME" } },
            { "RELAY_OVERCUR", new[] { "ALIASNAME" } },
            { "RELAY_MINMAX_ON_NODE", new[] { "ALIASNAME" } } };

        private static readonly Dictionary<string, string> Spaltentyp = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
            { "NETNAME", "TEXT(250)" }, { "NEPID", "INTEGER" }, { "NAME", "TEXT(80)" }, { "TYPE", "TEXT(80)" },
            { "ALIASNAME", "TEXT(250)" }, { "DESCRIPTION", "TEXT(250)" }, { "HYPERLINK", "TEXT(150)" },
            { "legacy", "TEXT(255)" }, { "panelname", "TEXT(255)" },
            { "NODE1", "TEXT(80)" }, { "NODE2", "TEXT(80)" }, { "NODE3", "TEXT(80)" }, { "NODE4", "TEXT(80)" },
            { "SW1", "SMALLINT" }, { "SW2", "SMALLINT" }, { "SW3", "SMALLINT" }, { "SW4", "SMALLINT" },
            { "VERSION", "TEXT(250)" }, { "PROJECT", "TEXT(250)" }, { "EXPORTNODES", "SMALLINT" },
            { "EXPORTGRAPHIC", "SMALLINT" }, { "EXPORT_TIME", "DATETIME" }, { "EXPORT_ONLY_FEEDED", "SMALLINT" },
            { "LFD", "INTEGER" }, { "NUMMER", "TEXT(50)" }, { "DATUM", "TEXT(20)" }, { "KUERZEL", "TEXT(50)" },
            { "PLANNAME", "TEXT(250)" } };

        private static readonly string[] TopoSpalten = { "NETNAME", "NEPID", "NAME", "TYPE", "NODE1", "NODE2", "NODE3", "NODE4", "SW1", "SW2", "SW3", "SW4" };
        private static readonly string[] InfoSpalten = { "NETNAME", "VERSION", "PROJECT", "EXPORTNODES", "EXPORTGRAPHIC", "EXPORT_TIME", "EXPORT_ONLY_FEEDED" };
        private static readonly string[] AendSpalten = { "NETNAME", "LFD", "NUMMER", "DATUM", "KUERZEL", "PLANNAME" };

        // Knoten und Schaltzustaende eines Elements, so wie NEPLAN sie in TOPOLOGY schreibt
        public static void Anschluesse(Projekt p, Element e, string[] knoten, int[] zustand)
        {
            for (int k = 0; k < 4; k++)
            {
                Element st = p.Station(e.Knoten[k]);
                knoten[k] = st == null ? "" : st.Name;
                zustand[k] = st == null ? -1 : e.Zustand[k];
            }
            Element l;
            if (e.Leitung != 0 && p.Elemente.TryGetValue(e.Leitung, out l))
            {
                // Schalter am Leitungsende: NODE1 = Leitung, NODE2 = Station, NODE3 = Art von NODE1
                knoten[0] = l.Name;
                zustand[0] = e.Zustand[0];
                knoten[2] = "LINE";
                zustand[2] = -1;
            }
        }

        public static string Wert(Element e, string spalte)
        {
            switch (spalte)
            {
                case "ALIASNAME": return e.Alias;
                case "DESCRIPTION": return e.Beschreibung;
                case "HYPERLINK": return e.Legacy;
                case "legacy": return e.Legacy;
                case "panelname": return e.Panelname;
                default: return "";
            }
        }

        public static Dictionary<string, int> Schreibe(Projekt p, string mdb, string netz)
        {
            try { return SchreibeIntern(p, mdb, netz); }
            finally { Mdb.Freigeben(); }
        }

        private static Dictionary<string, int> SchreibeIntern(Projekt p, string mdb, string netz)
        {
            var anzahl = new Dictionary<string, int>();
            using (OleDbConnection c = Mdb.Oeffne(mdb, true))
            {
                var spaltenJe = new Dictionary<string, List<string>>();
                foreach (Element e in p.Reihenfolge)
                {
                    string t;
                    if (!Tabelle.TryGetValue(e.Typ, out t) || spaltenJe.ContainsKey(t)) continue;
                    var l = new List<string> { "NETNAME", "NEPID", "NAME" };
                    l.AddRange(Felder[t]);
                    spaltenJe[t] = l;
                }
                spaltenJe["TOPOLOGY"] = new List<string>(TopoSpalten);
                spaltenJe["INFOTABLE"] = new List<string>(InfoSpalten);
                spaltenJe["A_AENDERUNGEN"] = new List<string>(AendSpalten);

                var laengen = new Dictionary<string, Dictionary<string, int>>();
                foreach (var kv in spaltenJe) laengen[kv.Key] = Sichere(c, kv.Key, kv.Value);

                using (OleDbTransaction tx = c.BeginTransaction())
                {
                    foreach (string t in spaltenJe.Keys)
                        Befehl(c, tx, "DELETE FROM [" + t + "] WHERE [NETNAME] = ?", new object[] { netz }).ExecuteNonQuery();

                    var knoten = new string[4];
                    var zustand = new int[4];
                    foreach (Element e in p.Reihenfolge)
                    {
                        string t;
                        if (Tabelle.TryGetValue(e.Typ, out t))
                        {
                            var w = new List<KeyValuePair<string, object>> { KV("NETNAME", netz), KV("NEPID", e.Nepid), KV("NAME", e.Name) };
                            foreach (string f in Felder[t]) w.Add(KV(f, Wert(e, f)));
                            Einfuegen(c, tx, t, laengen[t], w);
                            int n; anzahl.TryGetValue(t, out n); anzahl[t] = n + 1;
                        }
                        Anschluesse(p, e, knoten, zustand);
                        var z = new List<KeyValuePair<string, object>> { KV("NETNAME", netz), KV("NEPID", e.Nepid), KV("NAME", e.Name), KV("TYPE", e.Typ) };
                        for (int k = 0; k < 4; k++) z.Add(KV("NODE" + (k + 1), knoten[k]));
                        for (int k = 0; k < 4; k++) z.Add(KV("SW" + (k + 1), zustand[k] < 0 ? (object)DBNull.Value : (object)(short)zustand[k]));
                        Einfuegen(c, tx, "TOPOLOGY", laengen["TOPOLOGY"], z);
                    }

                    Einfuegen(c, tx, "INFOTABLE", laengen["INFOTABLE"], new List<KeyValuePair<string, object>> {
                        KV("NETNAME", netz), KV("VERSION", Leser.Version + " (ohne NEPLAN)"), KV("PROJECT", p.Datei),
                        KV("EXPORTNODES", (short)1), KV("EXPORTGRAPHIC", (short)0), KV("EXPORT_TIME", DateTime.Now),
                        KV("EXPORT_ONLY_FEEDED", (short)0) });

                    for (int i = 0; i < p.Aenderungen.Count; i++)
                        Einfuegen(c, tx, "A_AENDERUNGEN", laengen["A_AENDERUNGEN"], new List<KeyValuePair<string, object>> {
                            KV("NETNAME", netz), KV("LFD", i + 1), KV("NUMMER", p.Aenderungen[i].Nummer),
                            KV("DATUM", p.Aenderungen[i].Datum), KV("KUERZEL", p.Aenderungen[i].Kuerzel), KV("PLANNAME", p.Planname) });
                    tx.Commit();
                }
            }
            return anzahl;
        }

        private static KeyValuePair<string, object> KV(string k, object v) { return new KeyValuePair<string, object>(k, v); }

        private static int Laenge(string typ)
        {
            int a = typ.IndexOf('('), b = typ.IndexOf(')');
            return a > 0 && b > a ? int.Parse(typ.Substring(a + 1, b - a - 1)) : 0;
        }

        // Legt Tabelle und fehlende Spalten an, liefert die Textlaengen je Spalte
        private static Dictionary<string, int> Sichere(OleDbConnection c, string tabelle, List<string> spalten)
        {
            var laenge = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            DataTable t = c.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, tabelle, null });
            foreach (DataRow r in t.Rows)
                laenge[(string)r["COLUMN_NAME"]] = r["CHARACTER_MAXIMUM_LENGTH"] == DBNull.Value ? 0 : Convert.ToInt32(r["CHARACTER_MAXIMUM_LENGTH"]);
            if (laenge.Count == 0)
            {
                var sb = new StringBuilder();
                foreach (string s in spalten)
                {
                    if (sb.Length > 0) sb.Append(", ");
                    sb.Append("[" + s + "] " + Spaltentyp[s]);
                    laenge[s] = Laenge(Spaltentyp[s]);
                }
                Befehl(c, null, "CREATE TABLE [" + tabelle + "] (" + sb + ")", null).ExecuteNonQuery();
                return laenge;
            }
            foreach (string s in spalten)
            {
                if (laenge.ContainsKey(s)) continue;
                Befehl(c, null, "ALTER TABLE [" + tabelle + "] ADD COLUMN [" + s + "] " + Spaltentyp[s], null).ExecuteNonQuery();
                laenge[s] = Laenge(Spaltentyp[s]);
            }
            return laenge;
        }

        private static OleDbCommand Befehl(OleDbConnection c, OleDbTransaction tx, string sql, object[] werte)
        {
            var cmd = new OleDbCommand(sql, c);
            if (tx != null) cmd.Transaction = tx;
            if (werte != null) foreach (object v in werte) cmd.Parameters.Add(Parameter(v));
            return cmd;
        }

        private static OleDbParameter Parameter(object v)
        {
            var p = new OleDbParameter();
            if (v == null || v == DBNull.Value) { p.OleDbType = OleDbType.VarWChar; p.Value = DBNull.Value; }
            else if (v is DateTime) { p.OleDbType = OleDbType.Date; p.Value = v; }
            else if (v is short) { p.OleDbType = OleDbType.SmallInt; p.Value = v; }
            else if (v is int) { p.OleDbType = OleDbType.Integer; p.Value = v; }
            else { p.OleDbType = OleDbType.VarWChar; p.Value = v; }
            return p;
        }

        private static void Einfuegen(OleDbConnection c, OleDbTransaction tx, string tabelle, Dictionary<string, int> laengen,
                                      List<KeyValuePair<string, object>> werte)
        {
            var namen = new StringBuilder();
            var platz = new StringBuilder();
            var liste = new List<object>();
            foreach (var kv in werte)
            {
                if (namen.Length > 0) { namen.Append(", "); platz.Append(", "); }
                namen.Append("[" + kv.Key + "]");
                platz.Append("?");
                object v = kv.Value;
                string s = v as string;
                int max;
                // Leere Texte als NULL: von Jet angelegte Textspalten erlauben keine Leerzeichenkette
                if (s != null && s.Length == 0) v = DBNull.Value;
                else if (s != null && laengen.TryGetValue(kv.Key, out max) && max > 0 && s.Length > max) v = s.Substring(0, max);
                liste.Add(v);
            }
            Befehl(c, tx, "INSERT INTO [" + tabelle + "] (" + namen + ") VALUES (" + platz + ")", liste.ToArray()).ExecuteNonQuery();
        }
    }

    // ------------------------------------------------------------------ //
    // Vergleich mit einem echten NEPLAN-Export
    // ------------------------------------------------------------------ //
    public class Vergleich
    {
        public string Netz = "";
        public List<string> Zeilen = new List<string>();
        public int KernAbweichungen;          // Stationen, Leitungen, Trennschalter
        public int SonstigeAbweichungen;
        public bool NetzGefunden;
    }

    public static class Pruefer
    {
        private static readonly HashSet<string> Kern = new HashSet<string> { "BUSBAR-NODE", "LINE", "DISCSWITCH_2" };

        private static string T(object o) { return o == null || o == DBNull.Value ? "" : Convert.ToString(o).Trim(); }

        public static Vergleich Vergleiche(Projekt p, string neplanMdb, string netz)
        {
            try { return VergleicheIntern(p, neplanMdb, netz); }
            finally { Mdb.Freigeben(); }
        }

        private static Vergleich VergleicheIntern(Projekt p, string neplanMdb, string netz)
        {
            var v = new Vergleich { Netz = netz };
            using (OleDbConnection c = Mdb.Oeffne(neplanMdb, false))
            {
                var topo = Lies(c, "SELECT * FROM [TOPOLOGY] WHERE [NETNAME] = ?", netz);
                if (topo.Rows.Count == 0) { v.Zeilen.Add("Das Netz " + netz + " steht nicht in dieser Datenbank."); return v; }
                v.NetzGefunden = true;

                // Elemente und Anschluesse je Art
                var je = new SortedDictionary<string, int[]>();   // ok, abweichend, fehlt
                var bsp = new List<string>();
                var inMdb = new HashSet<int>();
                var knoten = new string[4];
                var zustand = new int[4];
                foreach (DataRow r in topo.Rows)
                {
                    string typ = T(r["TYPE"]);
                    int nep = Convert.ToInt32(r["NEPID"]);
                    inMdb.Add(nep);
                    int[] z;
                    if (!je.TryGetValue(typ, out z)) { z = new int[3]; je[typ] = z; }
                    Element e;
                    if (!p.Elemente.TryGetValue(nep, out e) || e.Name.Trim() != T(r["NAME"]) || e.Typ != typ)
                    {
                        z[2]++;
                        if (bsp.Count < 12) bsp.Add(typ + " " + T(r["NAME"]) + ": fehlt im Leser");
                        continue;
                    }
                    MdbSchreiber.Anschluesse(p, e, knoten, zustand);
                    string fehler = "";
                    for (int k = 0; k < 4; k++)
                    {
                        string sollK = T(r["NODE" + (k + 1)]), sollS = T(r["SW" + (k + 1)]);
                        string istS = zustand[k] < 0 ? "" : zustand[k].ToString();
                        if (sollK != knoten[k].Trim() || (sollK.Length > 0 && sollK != "LINE" && sollS != istS))
                            fehler += " NODE" + (k + 1) + " soll '" + sollK + "'/" + sollS + ", ist '" + knoten[k].Trim() + "'/" + istS;
                    }
                    if (fehler.Length == 0) z[0]++;
                    else
                    {
                        z[1]++;
                        if (bsp.Count < 12) bsp.Add(typ + " " + e.Name.Trim() + ":" + fehler);
                    }
                }
                int zusaetzlich = 0;
                foreach (Element e in p.Reihenfolge) if (!inMdb.Contains(e.Nepid)) zusaetzlich++;

                v.Zeilen.Add("Elemente und Anschlüsse (richtig / abweichend / fehlt):");
                foreach (var kv in je)
                {
                    v.Zeilen.Add(string.Format("  {0,-16} {1,5} {2,5} {3,5}", kv.Key, kv.Value[0], kv.Value[1], kv.Value[2]));
                    if (Kern.Contains(kv.Key)) v.KernAbweichungen += kv.Value[1] + kv.Value[2];
                    else v.SonstigeAbweichungen += kv.Value[1] + kv.Value[2];
                }
                if (zusaetzlich > 0) v.Zeilen.Add("  Im Leser, aber nicht im NEPLAN-Export: " + zusaetzlich);

                // Stationsfelder
                var spalten = Mdb.Spalten(c, "BUSBAR");
                var bus = Lies(c, "SELECT * FROM [BUSBAR] WHERE [NETNAME] = ?", netz);
                v.Zeilen.Add("Stationsfelder (richtig / alle):");
                foreach (string f in new[] { "ALIASNAME", "DESCRIPTION", "legacy", "panelname" })
                {
                    if (!spalten.Contains(f)) continue;
                    int ok = 0, alle = 0;
                    foreach (DataRow r in bus.Rows)
                    {
                        Element e;
                        if (!p.Elemente.TryGetValue(Convert.ToInt32(r["NEPID"]), out e)) continue;
                        alle++;
                        if (MdbSchreiber.Wert(e, f).Trim() == T(r[f])) ok++;
                        else
                        {
                            v.KernAbweichungen++;
                            if (bsp.Count < 12) bsp.Add("Station " + e.Name.Trim() + " " + f + ": soll '" + T(r[f]) + "', ist '" + MdbSchreiber.Wert(e, f).Trim() + "'");
                        }
                    }
                    v.Zeilen.Add(string.Format("  {0,-16} {1,5} / {2}", f, ok, alle));
                }
                if (bsp.Count > 0)
                {
                    v.Zeilen.Add("Beispiele für Abweichungen:");
                    foreach (string b in bsp) v.Zeilen.Add("  " + b);
                }
            }
            return v;
        }

        private static DataTable Lies(OleDbConnection c, string sql, string netz)
        {
            var cmd = new OleDbCommand(sql, c);
            cmd.Parameters.Add(new OleDbParameter { OleDbType = OleDbType.VarWChar, Value = netz });
            var t = new DataTable();
            new OleDbDataAdapter(cmd).Fill(t);
            return t;
        }
    }
}
