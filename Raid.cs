using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Collections;
using System.Linq;
using Newtonsoft.Json;

namespace ES_DKP_Utils
{
    public class Raid
    {

        #region Declarations
        private frmMain owner;

        private string _RaidName;
        public string RaidName
        {
            get { return _RaidName; }
            set
            {
                DataRow[] rs = dbRaid.Select("EventNameOrLoot='" + _RaidName + "' OR LootRaid='" + _RaidName + "'");
                foreach (DataRow r in rs)
                {
                    if (r["EventNameOrLoot"].Equals(_RaidName))
                    {
                        r["EventNameOrLoot"] = value;
                    }
                    if (r["LootRaid"].Equals(_RaidName))
                    {
                        r["LootRaid"] = value;
                    }
                }
                _RaidName = value;
            }
        }
        private DateTime _RaidDate;
        public DateTime RaidDate
        {
            get { return _RaidDate; }
            set
            {
                DataRow[] rs = dbRaid.Select("Date =#" + _RaidDate.Month + "/" + _RaidDate.Day + "/" + _RaidDate.Year + "#");
                foreach (DataRow r in rs)
                {
                    r["Date"] = value;
                }
                _RaidDate = value;
            }
        }
        private bool _AttendanceSynced;
        public bool AttendanceSynced
        {
            get
            {
                return _AttendanceSynced;
            }
            set
            {
                _AttendanceSynced = value;

            }
        }
        private bool _Loaded;
        public bool Loaded
        {
            get
            {
                return _Loaded;
            }
            set
            {
                _Loaded = value;

            }
        }
        private bool _DoubleTier;
        public bool DoubleTier
        {
            get
            {
                return _DoubleTier;
            }
            set
            {
                _DoubleTier = value;

            }
        }
        private System.Double _TotalDKP;
        public System.Double TotalDKP
        {
            get
            {
                return _TotalDKP;
            }
            set
            {
                _TotalDKP = value;

            }
        }
        private int _Attendees;
        public int Attendees
        {
            get
            {
                return _Attendees;
            }
            set
            {
                _Attendees = value;

            }
        }

        private DataTable dbRaid;
        private DataTable dbPeople;
        private DataTable dbAlts;

        private OleDbConnection dbConnect;
        private OleDbDataAdapter dkpDA;
        private OleDbDataAdapter altDA;
        private OleDbCommandBuilder cmdBld;
        private OleDbCommandBuilder altBld;

        private ArrayList ignore;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        #endregion

        #region Constuctor
        public Raid(frmMain owner)
        {
            log.Debug("Begin Method: Raid.Raid(frmMain) (" + owner.ToString() + ")");

            this.owner = owner;
            this.ignore = new ArrayList();
            DoubleTier = false;
            string connectionStr;
            connectionStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + owner.DBString;
            try { dbConnect = new OleDbConnection(connectionStr); }
            catch (Exception ex)
            {
                log.Error("Failed to create data connection: " + ex.Message);
                MessageBox.Show("Could not create data connection. \n(" + ex.Message + ")", "Error");
            }
            dkpDA = new OleDbDataAdapter();
            dbRaid = new DataTable("Raid");
            dbRaid.Columns.Add("Name", System.Type.GetType("System.String"));
            dbRaid.Columns.Add("Date", System.Type.GetType("System.DateTime"));
            dbRaid.Columns.Add("EventNameOrLoot", System.Type.GetType("System.String"));
            dbRaid.Columns.Add("PTS", System.Type.GetType("System.Double"));
            dbRaid.Columns.Add("Comment", System.Type.GetType("System.String"));
            dbRaid.Columns.Add("LootRaid", System.Type.GetType("System.String"));

            dbPeople = new DataTable("People");
            try
            {
                string query = "SELECT DISTINCT Name FROM DKS WHERE (Name<>\"zzzDKP Retired\")";
                dkpDA.SelectCommand = new OleDbCommand(query, dbConnect);
                dbConnect.Open();
                dkpDA.Fill(dbPeople);
            }
            catch (Exception ex)
            {
                log.Error("Failed to load people table: " + ex.Message);
                MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")", "Error");
            }
            finally { dbConnect.Close(); }

            log.Debug("End Method: Raid.Raid()");
        }
        #endregion

        #region Methods
        public void LoadRaid()
        {
            log.Debug("Begin Method: Raid.LoadRaid()");

            try
            {
                string query = "SELECT * FROM DKS WHERE (Date=#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "# AND (((EventNameOrLoot LIKE '" + RaidName + "%' AND LootRaid NOT LIKE '__%') OR LootRaid LIKE '" + RaidName + "%') AND EventNameOrLoot NOT LIKE '%1') )";
                dkpDA.SelectCommand = new OleDbCommand(query, dbConnect);
                dbConnect.Open();
                dkpDA.Fill(dbRaid);
                cmdBld = new OleDbCommandBuilder(dkpDA);
                Loaded = true;

                DataRow[] tiercheck = dbRaid.Select("EventNameOrLoot LIKE '%0' OR LootRaid LIKE '%0'");
                if (tiercheck.Length > 0)
                {
                    DoubleTier = true;
                    foreach (DataRow r in tiercheck)
                    {
                        bool isloot = (double)r.ItemArray[3] < 0;
                        if (isloot)
                        {
                            string s = (string)r.ItemArray[5];
                            s = s.TrimEnd(new char[] { '0' });
                            r.ItemArray = new object[] {
                                                           r.ItemArray[0], r.ItemArray[1],
                                                           r.ItemArray[2], r.ItemArray[3],
                                                           r.ItemArray[4], s };
                        }
                        else
                        {
                            string s = (string)r.ItemArray[2];
                            s = s.TrimEnd(new char[] { '0' });
                            r.ItemArray = new object[] {
                                                           r.ItemArray[0], r.ItemArray[1],
                                                           s, r.ItemArray[3],
                                                           r.ItemArray[4], r.ItemArray[5] };
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                log.Error("Failed to load raid: " + ex.Message);
                MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")", "Error");
            }
            finally { dbConnect.Close(); }

            log.Debug("End Method: Raid.LoadRaid()");
        }

        public void ShowAttendees()
        {
            log.Debug("Begin Method: Raid.ShowAttendees()");

            DataRow[] attendees = dbRaid.Select("(PTS>=0)");
            DataTable newTable = rowsToTable(dbRaid, attendees);
            TableViewer table = new TableViewer(owner, newTable, new TableViewer.UpdateDel(UpdateAttendees));
            table.Show();

            log.Debug("End Method: Raid.ShowAttendees()");
        }
        public void UpdateAttendees(DataTable dt)
        {
            log.Debug("Begin Method: Raid.UpdateAttendees(DataTable) (" + dt.ToString() + ")");

            foreach (DataRow r in dbRaid.Select("(PTS>=0)")) { dbRaid.Rows.Remove(r); }
            foreach (DataRow r in dt.Rows) { dbRaid.Rows.Add(r.ItemArray); }

            log.Debug("End Method: Raid.UpdateAttendees()");
        }

        public void ShowLoot()
        {
            log.Debug("Begin Method: Raid.ShowLoot()");

            DataRow[] loot = dbRaid.Select("(PTS<0)");
            DataTable newTable = rowsToTable(dbRaid, loot);
            TableViewer table = new TableViewer(owner, newTable, new TableViewer.UpdateDel(UpdateLoot));
            table.Show();

            log.Debug("End Method: Raid.ShowLoot()");
        }

        public void UpdateLoot(DataTable dt)
        {
            log.Debug("Begin Method: Raid.UpdateLoot(DataTable) (" + dt.ToString() + ")");

            foreach (DataRow r in dbRaid.Select("(PTS<0)")) { dbRaid.Rows.Remove(r); }
            foreach (DataRow r in dt.Rows) { dbRaid.Rows.Add(r.ItemArray); }

            log.Debug("End Method: Raid.UpdateLoot()");
        }

        public void InputPerson(string s)
        {
            log.Debug("Begin Method: Raid.InputPerson(string) (" + s.ToString() + ")");

            DataRow[] rows = dbPeople.Select("Name='" + s + "'");

            if (rows.Length > 0)
            {
                DataRow[] rs = dbRaid.Select("Name='" + s + "' AND PTS>=0");
                if (rs.Length == 0)
                {
                    dbRaid.Rows.Add(new object[] { s, RaidDate, RaidName, 0, null, null });
                    log.Debug("Added person: " + s);
                }
            }
            else
            {
                string t = GetAlt(s);
                if ((t != null && t.Length != 0))
                {
                    if (dbRaid.Select("Name='" + t + "' AND PTS>=0").Length == 0)
                    {
                        DataRow[] rs = dbRaid.Select("Name='" + t + "' AND PTS>=0");
                        if (rs.Length == 0)
                        {
                            dbRaid.Rows.Add(new object[] { t, RaidDate, RaidName, 0, null, null });
                            log.Debug("Added person: " + t);
                        }

                    }

                }
                else
                {
                    if (!ignore.Contains(t))
                    {
                        AppAltDialog alt = new AppAltDialog(owner, s, dbPeople);
                        t = alt.GetName();
                        if ((t != null && t.Length != 0))
                        {
                            DataRow[] rs = dbRaid.Select("Name='" + t + "' AND PTS>=0");
                            if (rs.Length == 0)
                            {
                                dbRaid.Rows.Add(new object[] { t, RaidDate, RaidName, 0, null, null });
                                log.Debug("Added person: " + t);
                            }
                            if (t != s) dbAlts.Rows.Add(new object[] { s, t });
                            altDA.Update(dbAlts);
                        }
                        else { ignore.Add(t); }
                    }
                }
            }
        }

        public void ParseRaidDump(string logFile)
        {
            log.Debug("Begin Method: Raid.ParseRaidDump(string) (" + logFile.ToString() + ")");

            FileStream input;
            StreamReader sr;
            string line;
            char[] delimiters = new char[] { '\t' };
            ArrayList people = new ArrayList();

            try
            {
                input = new FileStream(logFile, FileMode.Open, FileAccess.Read);
                sr = new StreamReader(input);
            }
            catch (Exception ex)
            {
                log.Error("Failed to open raid dump file to parse: " + ex.Message);
                MessageBox.Show("File IO Error.\n(" + ex.Message + ")");
                return;
            }

            do
            {
                line = sr.ReadLine();

                if (line == null)
                {
                    log.Debug("Line in file empty, skipping");
                    break;
                }

                string[] parts = line.Split(delimiters);

                people.Add(parts[1]);
            } while (line != null);

            sr.Close();
            input.Close();

            if (people.Count > 0)
            {
                MessageBox.Show("Found " + people.Count + " raiders including " + people[0] + ".");
            }

            foreach (string s in people)
            {
                DataRow[] rows = dbPeople.Select("Name='" + s + "'");

                if (rows.Length > 0)
                {
                    DataRow[] rs = dbRaid.Select("Name='" + s + "' AND PTS>=0");
                    if (rs.Length == 0)
                    {
                        dbRaid.Rows.Add(new object[] { s, RaidDate, RaidName, 0, null, null });
                        log.Debug("Added attendance row for " + s);
                    }
                }
                else
                {
                    string t = GetAlt(s);
                    if ((t != null && t.Length != 0))
                    {
                        if (!people.Contains(t))
                        {
                            DataRow[] rs = dbRaid.Select("Name='" + t + "' AND PTS>=0");
                            if (rs.Length == 0)
                            {
                                dbRaid.Rows.Add(new object[] { t, RaidDate, RaidName, 0, null, null });
                                log.Debug("Added attendance row for " + t + " (Alt: " + s + ")");
                            }
                        }

                    }
                    else
                    {
                        AppAltDialog alt = new AppAltDialog(owner, s, dbPeople);
                        t = alt.GetName();
                        if ((t != null && t.Length != 0))
                        {
                            DataRow[] rs = dbRaid.Select("Name='" + t + "' AND PTS>=0");
                            if (rs.Length == 0)
                            {
                                dbRaid.Rows.Add(new object[] { t, RaidDate, RaidName, 0, null, null });
                                log.Debug("Added attendance row for " + t);
                            }
                            if (t != s) dbAlts.Rows.Add(new object[] { s, t });
                            altDA.Update(dbAlts);
                        }
                    }
                }

            }

            log.Debug("End Method: Raid.ParseRaidDump()");
        }

        public void ParseAttendance(string logFile, string zo)
        {
            log.Debug("Begin Method: Raid.ParseAttendance(string,string) (" + logFile.ToString() + "," + zo.ToString() + ")");

            string[] zones = zo.Split(new char[] { ' ' });
            string line;
            FileStream input;
            StreamReader sr;
            Regex regex = new Regex(@"\[.*\] \[.*\] (?<name>\S+).*<(?<guild>.*)> ZONE: (?<zone>.*)$");
            Match m;
            ArrayList people = new ArrayList();

            foreach (DataRow z in dbRaid.Rows) people.Add(z.ItemArray[0]);

            try
            {
                input = new FileStream(logFile, FileMode.Open, FileAccess.Read);
                sr = new StreamReader(input);
            }
            catch (Exception ex)
            {
                log.Error("Failed to open log to parse: " + ex.Message);
                MessageBox.Show("File IO Error.\n(" + ex.Message + ")");
                return;
            }

            do
            {
                line = sr.ReadLine();

                if (line != null)
                {
                    m = regex.Match(line);
                    if (m.Success)
                    {
                        string _guild = m.Groups["guild"].ToString();
                        string _zone = m.Groups["zone"].ToString();
                        string _name = m.Groups["name"].ToString();

                        if (owner.GuildNames.Contains(_guild) && zo.Contains(_zone) && !people.Contains(_name))
                        {
                            log.Debug(_name + " <" + _guild + "> is in zone " + _zone + ", adding to people array.");
                            people.Add(_name);
                        }
                    }
                }
            } while (line != null);

            sr.Close();
            input.Close();

            people.Sort();

            foreach (string s in people)
            {
                DataRow[] rows = dbPeople.Select("Name='" + s + "'");

                if (rows.Length > 0)
                {
                    DataRow[] rs = dbRaid.Select("Name='" + s + "' AND PTS>=0");
                    if (rs.Length == 0)
                    {
                        dbRaid.Rows.Add(new object[] { s, RaidDate, RaidName, 0, null, null });
                        log.Debug("Added attendance row for " + s);
                    }
                }
                else
                {
                    string t = GetAlt(s);
                    if ((t != null && t.Length != 0))
                    {
                        if (!people.Contains(t))
                        {
                            DataRow[] rs = dbRaid.Select("Name='" + t + "' AND PTS>=0");
                            if (rs.Length == 0)
                            {
                                dbRaid.Rows.Add(new object[] { t, RaidDate, RaidName, 0, null, null });
                                log.Debug("Added attendance row for " + t + " (Alt: " + s + ")");
                            }
                        }

                    }
                    else
                    {
                        AppAltDialog alt = new AppAltDialog(owner, s, dbPeople);
                        t = alt.GetName();
                        if ((t != null && t.Length != 0))
                        {
                            DataRow[] rs = dbRaid.Select("Name='" + t + "' AND PTS>=0");
                            if (rs.Length == 0)
                            {
                                dbRaid.Rows.Add(new object[] { t, RaidDate, RaidName, 0, null, null });
                                log.Debug("Added attendance row for " + t);
                            }
                            if (t != s) dbAlts.Rows.Add(new object[] { s, t });
                            altDA.Update(dbAlts);
                        }
                    }
                }

            }
            log.Debug("Begin Method: Raid.ParseAttendance()");
        }

        public string GetAlt(string s)
        {
            log.Debug("Begin Method: Raid.GetAlt(string) (" + s.ToString() + ")");

            dbAlts = new DataTable("Alts");
            string query = "SELECT * FROM ALTS";
            try
            {
                altDA = new OleDbDataAdapter(query, dbConnect);
                dbConnect.Open();
                altDA.Fill(dbAlts);
                altBld = new OleDbCommandBuilder(altDA);
            }
            catch (Exception ex)
            {
                log.Error("Failed to load alts table: " + ex.Message);
                MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")", "Error");
            }
            finally { dbConnect.Close(); }

            DataRow[] rows = dbAlts.Select("AltName='" + s + "'");

            if (rows.Length > 0)
            {
                log.Debug("End Method: Raid.GetAlt(), returning " + (string)rows[0].ItemArray[1]); ;
                return (string)rows[0].ItemArray[1];
            }

            log.Debug("End Method: Raid.GetAlt(), returning ");
            return "";

        }

        public string[] ParseLine(string line)
        {
            /* Whatever is using this function is expecting a string array to come back with these elements:
             *   0: Raider's Name
             *   1: Zone
             *   
             * I've added guild as paramater [2], in case we want to use it in guild name matching.
             */
            log.Debug("Begin Method: Raid.ParseLine(string) (" + line.ToString() + ")");

            Regex r = new Regex(@"\[.*\] \[.*\] (?<name>\S+).*<(?<guild>.*)> ZONE: (?<zone>.*)$");
            Match m = r.Match(line.Trim());

            string[] results = { m.Groups["name"].ToString(),
                                 m.Groups["zone"].ToString(),
                                 m.Groups["guild"].ToString() };

            log.Debug("End Method: Raid.ParseLine(), returning {" + results[0] + "," + results[1] + "}");
            return results;
        }

        public void AddRowToLocalTable(object[] o)
        {
            log.Debug("Begin Method: Raid.AddRowToLocalTable(object[]) (" + o.ToString() + ")");
            string s = "Creating new row: ";
            try
            {
                foreach (object k in o)
                {
                    if (k != null)
                        s += k.ToString() + ",";
                    else s += "null,";
                }
                log.Debug(s.TrimEnd(new char[] { ',' }));
            }
            catch (Exception ex) {
                log.Error("Error adding row to local table: " + ex.Message);
            }


            DataRow r = dbRaid.NewRow();
            r.ItemArray = o;
            dbRaid.Rows.Add(r);

            log.Debug("End Method: Raid.AddRowToLocalTable()");
        }

        public void DivideDKP()
        {
            log.Debug("Begin Method: Raid.DivideDKP()");

            double dkp = 0.0;
            DataRow[] rows = dbRaid.Select("PTS<0 AND LootRaid='" + RaidName + "'");
            foreach (DataRow r in rows) { dkp += Math.Abs((double)r.ItemArray[3]); }
            log.Info("Total dkp to distribute: " + dkp);

            rows = dbRaid.Select("PTS>=0 AND EventNameOrLoot='" + RaidName + "'");
            if (rows.Length == 0) rows = dbRaid.Select("PTS>=0 AND EventNameOrLoot='" + RaidName + "0'");
            if (rows.Length == 0) return;

            Attendees = rows.Length;
            log.Info("Total attendees: " + Attendees);

            TotalDKP = dkp;
            dkp *= (1 - owner.DKPTax);
            dkp /= rows.Length;

            log.Info("DKP per person: " + dkp);

            foreach (DataRow r in rows)
            {
                r.ItemArray = new object[] { r.ItemArray[0], r.ItemArray[1], r.ItemArray[2], dkp, r.ItemArray[4], r.ItemArray[5] };
            }

            log.Debug("End Method: Raid.DivideDKP()");
        }

        public void SyncData()
        {
            int k = 0;
            log.Debug("Begin Method: Raid.SyncData()");

            owner.StatusMessage = "Saving raid...";
            owner.PBMin = 0;
            owner.PBMax = dbRaid.Rows.Count + 1;
            owner.PBVal = 0;

            string query = "UPDATE DKS SET Comment='DeleteMe' WHERE Date=#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "# AND ( (EventNameOrLoot LIKE '" + RaidName + "%' AND PTS>=0)OR (LootRaid LIKE '" + RaidName + "%' AND PTS<0))";
            OleDbCommand updateCommand = null;
            OleDbCommand insertCommand = null;
            OleDbCommand deleteCommand = null;
            try
            {
                updateCommand = new OleDbCommand(query, dbConnect);
                dbConnect.Open();
                updateCommand.ExecuteNonQuery();
                log.Info("Updated " + k + " rows, flagged for deletion in DKS");

                query = "UPDATE EventLog SET DKPValue=-1 WHERE EventDate=#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "# AND ( EventName LIKE '" + RaidName + "%' )";
                updateCommand = new OleDbCommand(query, dbConnect);
                k = updateCommand.ExecuteNonQuery();
                log.Info("Updated " + k + " rows, flagged for deletion in EventLog");
            }
            catch (Exception ex)
            {
                log.Error("Failed to load either DKS or EventLog: " + ex.Message);
                MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")", "Error");
            }
            finally { dbConnect.Close(); }

            DataRow[] ppl = dbRaid.Select("PTS>=0");
            DataRow[] loot = dbRaid.Select("PTS<0");

            string lootstr;

            if (!DoubleTier)
            {
                try
                {
                    dbConnect.Open();
                    foreach (DataRow r in ppl)
                    {
                        query = "UPDATE DKS SET EventNameOrLoot=\"" + r.ItemArray[2] + "\", PTS=" + r.ItemArray[3] + ", Comment=\"" + r.ItemArray[4] + "\", LootRaid=\"" + r.ItemArray[5] + "\" WHERE (Name=\"" + r.ItemArray[0] + "\" AND PTS>=0 AND Date=#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "# AND EventNameOrLoot=\"" + RaidName + "\")";
                        updateCommand = new OleDbCommand(query, dbConnect);
                        if ((k = updateCommand.ExecuteNonQuery()) == 0)
                        {
                            query = "INSERT INTO DKS VALUES (\"" + r.ItemArray[0] + "\",#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "#,\"" + r.ItemArray[2] + "\"," + r.ItemArray[3] + ",\"" + r.ItemArray[4] + "\",\"" + r.ItemArray[5] + "\")";
                            insertCommand = new OleDbCommand(query, dbConnect);
                            k = insertCommand.ExecuteNonQuery();
                        }
                        owner.PBVal++;
                    }
                    foreach (DataRow r in loot)
                    {
                        lootstr = (string)r.ItemArray[2];
                        query = "INSERT INTO DKS VALUES (\"" + r.ItemArray[0] + "\",#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "#,\"" + lootstr + "\"," + r.ItemArray[3] + ",\"" + r.ItemArray[4] + "\",\"" + r.ItemArray[5] + "\")";
                        insertCommand = new OleDbCommand(query, dbConnect);
                        insertCommand.ExecuteNonQuery();

                        owner.PBVal++;
                    }
                    DataTable temp = new DataTable();
                    query = "SELECT EventName FROM EventLog WHERE EventDate=#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "# AND EventName NOT LIKE '%1'";
                    OleDbDataAdapter tempda = new OleDbDataAdapter(query, dbConnect);
                    tempda.Fill(temp);

                    query = "UPDATE EventLog ";
                    query += "SET Points=" + (TotalDKP / Attendees) * (1 - owner.DKPTax) + ", DKPValue=" + TotalDKP + " ";
                    query += "WHERE (EventDate=#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "# AND EventName=\"" + RaidName + "\")";

                    updateCommand = new OleDbCommand(query, dbConnect);
                    if (updateCommand.ExecuteNonQuery() == 0)
                    {
                        query = "INSERT INTO EventLog Values (#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "#,\"" + RaidName + "\"," + (TotalDKP / Attendees) * (1 - owner.DKPTax) + "," + TotalDKP + "," + (int)((int)temp.Rows.Count + 1) + ")";
                        insertCommand = new OleDbCommand(query, dbConnect);
                        insertCommand.ExecuteNonQuery();
                    }
                    owner.PBVal++;

                    query = "DELETE FROM DKS WHERE Comment='DeleteMe'";
                    deleteCommand = new OleDbCommand(query, dbConnect);
                    deleteCommand.ExecuteNonQuery();
                    query = "DELETE FROM EventLog WHERE DKPValue=-1";
                    deleteCommand = new OleDbCommand(query, dbConnect);
                    deleteCommand.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    log.Error("Failed to update DKS or EventLog: " + ex.Message);
                    MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")", "Error");
                }
                finally { dbConnect.Close(); }
            }
            else
            {
                try
                {
                    dbConnect.Open();
                    foreach (DataRow r in ppl)
                    {
                        query = "UPDATE DKS SET EventNameOrLoot=\"" + r.ItemArray[2] + "0\", PTS=" + r.ItemArray[3] + ", Comment=\"" + r.ItemArray[4] + "\", LootRaid=\"" + r.ItemArray[5] + "\" WHERE (Name=\"" + r.ItemArray[0] + "\" AND PTS>=0 AND Date=#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "# AND EventNameOrLoot=\"" + RaidName + "0\")";
                        updateCommand = new OleDbCommand(query, dbConnect);
                        if (updateCommand.ExecuteNonQuery() == 0)
                        {
                            query = "INSERT INTO DKS VALUES (\"" + r.ItemArray[0] + "\",#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "#,\"" + r.ItemArray[2] + "0\"," + r.ItemArray[3] + ",\"" + r.ItemArray[4] + "\",\"" + r.ItemArray[5] + "\")";
                            insertCommand = new OleDbCommand(query, dbConnect);
                            insertCommand.ExecuteNonQuery();
                        }

                        query = "UPDATE DKS SET EventNameOrLoot=\"" + r.ItemArray[2] + "1\", PTS=0, Comment=\"" + r.ItemArray[4] + "\", LootRaid=\"" + r.ItemArray[5] + "\" WHERE (Name=\"" + r.ItemArray[0] + "\" AND PTS>=0 AND Date=#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "# AND EventNameOrLoot=\"" + RaidName + "1\")";
                        updateCommand = new OleDbCommand(query, dbConnect);
                        if (updateCommand.ExecuteNonQuery() == 0)
                        {
                            query = "INSERT INTO DKS VALUES (\"" + r.ItemArray[0] + "\",#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "#,\"" + r.ItemArray[2] + "1\",0,\"" + r.ItemArray[4] + "\",\"" + r.ItemArray[5] + "\")";
                            insertCommand = new OleDbCommand(query, dbConnect);
                            insertCommand.ExecuteNonQuery();
                        }

                        owner.PBVal++;
                    }
                    foreach (DataRow r in loot)
                    {
                        lootstr = (string)r.ItemArray[2];
                        query = "UPDATE DKS ";
                        query += "SET PTS=" + r.ItemArray[3] + ", Comment=\"" + r.ItemArray[4] + "\" ";
                        query += "WHERE (Name=\"" + r.ItemArray[0] + "\" AND PTS<0 AND Date=#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "# AND LootRaid=\"" + RaidName + "0\" AND EventNameOrLoot=\"" + lootstr + "\")";
                        updateCommand = new OleDbCommand(query, dbConnect);

                        if (updateCommand.ExecuteNonQuery() == 0)
                        {
                            query = "INSERT INTO DKS VALUES (\"" + r.ItemArray[0] + "\",#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "#,\"" + lootstr + "\"," + r.ItemArray[3] + ",\"" + r.ItemArray[4] + "\",\"" + r.ItemArray[5] + "0\")";
                            insertCommand = new OleDbCommand(query, dbConnect);
                            insertCommand.ExecuteNonQuery();
                        }

                        owner.PBVal++;
                    }
                    DataTable temp = new DataTable();
                    query = "SELECT EventName FROM EventLog WHERE EventDate=#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "# AND EventName NOT LIKE '%1'";
                    OleDbDataAdapter tempda = new OleDbDataAdapter(query, dbConnect);
                    tempda.Fill(temp);

                    query = "UPDATE EventLog ";
                    query += "SET Points=" + (TotalDKP / Attendees) * (1 - owner.DKPTax) + ", DKPValue=" + TotalDKP + " ";
                    query += "WHERE (EventDate=#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "# AND EventName=\"" + RaidName + "0\")";

                    updateCommand = new OleDbCommand(query, dbConnect);
                    if (updateCommand.ExecuteNonQuery() == 0)
                    {
                        query = "INSERT INTO EventLog Values (#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "#,\"" + RaidName + "0\"," + (TotalDKP / Attendees) * (1 - owner.DKPTax) + "," + TotalDKP + "," + (int)((int)temp.Rows.Count + 1) + ")";
                        insertCommand = new OleDbCommand(query, dbConnect);
                        insertCommand.ExecuteNonQuery();
                    }

                    query = "UPDATE EventLog ";
                    query += "SET Points=0, DKPValue=0 ";
                    query += "WHERE (EventDate=#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "# AND EventName=\"" + RaidName + "1\")";

                    updateCommand = new OleDbCommand(query, dbConnect);
                    if (updateCommand.ExecuteNonQuery() == 0)
                    {
                        query = "INSERT INTO EventLog Values (#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "#,\"" + RaidName + "1\"," + 0 + "," + 0 + "," + (int)((int)temp.Rows.Count + 1) + ")";
                        insertCommand = new OleDbCommand(query, dbConnect);
                        insertCommand.ExecuteNonQuery();
                    }

                    owner.PBVal++;

                    query = "DELETE FROM DKS WHERE Comment='DeleteMe'";
                    deleteCommand = new OleDbCommand(query, dbConnect);
                    deleteCommand.ExecuteNonQuery();
                    query = "DELETE FROM EventLog WHERE DKPValue=-1";
                    deleteCommand = new OleDbCommand(query, dbConnect);
                    deleteCommand.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    log.Error("Failed to update DKS or EventLog: " + ex.Message);
                    MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")", "Error");
                }
                finally { dbConnect.Close(); }
            }

            if (true)
            {
                MessageBox.Show(JsonConvert.SerializeObject(this));
            }

            log.Debug("End Method: Raid.SyncRaid()");
        }

        public void FigureTiers()
        {
            log.Debug("Begin Method: Raid.FigureTiers()");

            string query = "SELECT DISTINCT EventDate FROM EventLog WHERE EventDate <= #" + DateTime.Now.Month + "/" + DateTime.Now.Day + "/" + DateTime.Now.Year + "#";
            DataTable dates = null;
            DataTable nameCount = null;
            try
            {
                dates = new DataTable("Date");
                dkpDA = new OleDbDataAdapter(query, dbConnect);
                dbConnect.Open();
                dkpDA.Fill(dates);
            }
            catch (Exception ex)
            {
                log.Error("Failed to fill dates table: " + ex.Message);
                MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")", "Error");
            }
            finally { dbConnect.Close(); }

            DateTime tierwindow = (DateTime)dates.Rows[dates.Rows.Count - 20].ItemArray[0];
            int totalraids = 0;
            try
            {
                OleDbCommand deleteCmd = new OleDbCommand("DELETE * FROM NamesTiers", dbConnect);
                dbConnect.Open();
                deleteCmd.ExecuteNonQuery();
                nameCount = new DataTable("DKS");
                query = "SELECT DKS.Name, Count(DKS.Name) as NameCount FROM DKS WHERE DKS.Date>=#" + tierwindow.Month + "/" + tierwindow.Day + "/" + tierwindow.Year + "# AND DKS.PTS>=0 GROUP BY DKS.Name";
                dkpDA = new OleDbDataAdapter(query, dbConnect);
                dkpDA.Fill(nameCount);
                query = "SELECT Count(EventLog.EventName) as TotalRaids FROM EventLog WHERE EventLog.EventDate>=#" + tierwindow.Month + "/" + tierwindow.Day + "/" + tierwindow.Year + "#";
                dkpDA = new OleDbDataAdapter(query, dbConnect);
                DataTable t = new DataTable("temp");
                dkpDA.Fill(t);
                totalraids = (int)t.Rows[0].ItemArray[0];
            }
            catch (Exception ex)
            {
                log.Error("Failed to fill attendance count or raid count table: " + ex.Message);
                MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")", "Error");
            }
            finally { dbConnect.Close(); }

            owner.PBMin = 0;
            owner.PBMax = nameCount.Rows.Count + nameCount.Rows.Count / 10;
            owner.PBVal = nameCount.Rows.Count / 10;

            DataRow[] rows = nameCount.Select("Name<>'zzzDKP Retired'");
            foreach (DataRow r in rows)
            {
                int numraids = (int)r.ItemArray[1];
                char c;
                double d;
                string q;
                d = (double)numraids / totalraids;
                if (d >= owner.TierAPct)
                    c = 'A';
                else if (d >= owner.TierBPct)
                    c = 'B';
                else if (d >= owner.TierCPct)
                    c = 'C';
                else
                    c = 'E';
                q = "INSERT INTO NamesTiers VALUES (\"" + r.ItemArray[0] + "\",\"" + c + "\"," + numraids + "," + d * 100 + ")";
                try
                {
                    OleDbCommand dbc = new OleDbCommand(q, dbConnect);
                    dbConnect.Open();
                    dbc.ExecuteNonQuery();
                    owner.PBVal++;
                }
                catch (Exception ex)
                {
                    log.Error("Failed to write NamesTiers row: " + ex.Message);
                    MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")", "Error");
                }
                finally { dbConnect.Close(); }

            }

            log.Debug("End Method: Raid.FigureTiers()");
        }

        public void AddNameClass(string s, string t)
        {
            log.Debug("Begin Method: Raid.AddNameClass(string,string) (" + s.ToString() + "," + t.ToString() + ")");

            string query = "INSERT INTO NamesClassesRemix VALUES ('" + s + "','" + t + "')";
            try
            {
                OleDbCommand nameclass = new OleDbCommand(query, dbConnect);
                dbConnect.Open();
                nameclass.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                log.Error("Failed to insert NamesClasses row: " + ex.Message);
                MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")", "Error");
            }
            finally { dbConnect.Close(); }

            log.Debug("End Method: Raid.AddNameClass()");
        }

        public void AddPerson(string s)
        {
            log.Debug("Begin Method: Raid.AddPerson(string) (" + s.ToString() + ")");

            if (Present().Select("Name='" + s + "' AND PTS>=0").Length == 0)
            {
                object[] o = new object[] { s, RaidDate, RaidName, 0, "", "" };
                AddRowToLocalTable(o);
                DivideDKP();
            }

            log.Debug("End Method: Raid.AddPerson()");
        }

        public void RemovePerson(string s)
        {
            log.Debug("Begin Method: Raid.RemovePerson(string) (" + s.ToString() + ")");

            DataRow[] rows = dbRaid.Select("Name='" + s + "' AND PTS>=0");
            try { dbRaid.Rows.Remove(rows[0]); }
            catch (Exception ex)
            {
                log.Error("Failed to remove " + s + " from local table: " + ex.Message);
            }
            DivideDKP();

            log.Debug("End Method: Raid.RemovePerson()");
        }

        public DataTable NotPresent()
        {
            log.Debug("Begin Method: Raid.NotPresent()");

            DataTable p = dbPeople.Copy();
            DataRow[] deleteMe = new DataRow[p.Rows.Count];
            int i = 0;
            int j = 0;
            foreach (DataRow r in p.Rows)
            {
                if (dbRaid.Select("Name='" + r.ItemArray[0] + "'").Length > 0)
                {
                    deleteMe[i] = r;
                    i++;
                }
            }

            for (j = 0; j < i; j++) p.Rows.Remove(deleteMe[j]);

            log.Debug("End Method: Raid.NotPresent(), returning " + p.ToString());
            return p;
        }

        public DataTable Present()
        {
            log.Debug("Begin Method: Raid.Present()");

            DataTable p = dbRaid.Copy();

            log.Debug("End Method: Raid.Present(),  returning " + p);
            return p;
        }

        public static DataTable rowsToTable(DataTable baseTable, DataRow[] rows)
        {
            DataTable t = baseTable.Clone();
            foreach (DataRow r in rows)
            {
                t.ImportRow(r);
            }
            return t;
        }

        public static void DailyReport(frmMain owner, DateTime t)
        {
            OleDbConnection dbCon = null;
            DataTable dtTotal = new DataTable("DKSALL");
            DataTable dtDKS = new DataTable("DKS");
            DataTable dtEL = new DataTable("EventLog");
            DataTable dtNT = new DataTable("NamesTiers");
            DataTable dtAllEvents = new DataTable("NTALL");
            DataTable dtFinal = new DataTable("FINAL");

            FileStream fs = null;
            FileStream fs2 = null;
            StreamWriter sw = null;
            StreamWriter sw2 = null;
            int i = 1;
            double dkp = 0.0;
            int ppl = 1;
            bool twotier = false;
            double earned;
            double totalnetchange = 0.0;
            double totaldkp = 0.0;

            owner.PBMin = 0;
            owner.PBMax = 110;
            owner.PBVal = 0;

            string connectionStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + owner.DBString;
            try { dbCon = new OleDbConnection(connectionStr); }
            catch (Exception ex)
            {
                MessageBox.Show("Could not create data connection. \n(" + ex.Message + ")", "Error");
            }

            try
            {
                string query = "SELECT * FROM DKS WHERE Date=#" + t.Month + "/" + t.Day + "/" + t.Year + "# ORDER BY Name";
                OleDbDataAdapter da = new OleDbDataAdapter(query, dbCon);
                dbCon.Open();
                da.Fill(dtDKS);
                query = "SELECT * FROM EventLog WHERE EventDate=#" + t.Month + "/" + t.Day + "/" + t.Year + "# ORDER BY RaidNumber ASC";
                da = new OleDbDataAdapter(query, dbCon);
                da.Fill(dtEL);
                owner.PBVal += 10;
                query = "SELECT * FROM NamesTiers";
                da = new OleDbDataAdapter(query, dbCon);
                da.Fill(dtNT);
                owner.PBVal += 10;
                query = "SELECT Name, Sum(PTS) AS Total FROM DKS GROUP BY Name HAVING Name<>'zzzDKP Retired'";
                da = new OleDbDataAdapter(query, dbCon);
                da.Fill(dtTotal);
                owner.PBVal += 10;
                DataTable temp = new DataTable();
                query = "SELECT DISTINCT TOP " + owner.RaidDaysWindow + " EventDate FROM EventLog ORDER BY EventDate DESC";
                da = new OleDbDataAdapter(query, dbCon);
                da.Fill(temp);
                owner.PBVal += 10;
                DateTime tempdate = (DateTime)temp.Rows[owner.RaidDaysWindow - 1]["EventDate"];
                query = "SELECT * FROM EventLog WHERE EventDate>=#" + tempdate.Month + "/" + tempdate.Day + "/" + tempdate.Year + "# ORDER BY EventDate ASC, RaidNumber ASC";
                da = new OleDbDataAdapter(query, dbCon);
                da.Fill(dtAllEvents);
                owner.PBVal += 10;
                query = "SELECT DKS.Name, Sum(DKS.PTS) AS Total, Max(DKS.Date) AS LastRaid, NamesTiers.Tier, NamesTiers.NumRaids, NamesTiers.TPercent ";
                query += "FROM DKS INNER JOIN NamesTiers ON DKS.Name=NamesTiers.Name GROUP BY DKS.Name, NamesTiers.Tier, NamesTiers.TPercent, NamesTiers.NumRaids";
                da = new OleDbDataAdapter(query, dbCon);
                da.Fill(dtFinal);
                owner.PBVal += 10;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")", "Error");
            }
            finally { dbCon.Close(); }

            try
            {
                // Terrible hack while I convert to BBCode output
                fs = new FileStream(owner.OutputDirectory + "DKPReport.txt", FileMode.Create);
                sw = new StreamWriter(fs);
                fs2 = new FileStream(owner.OutputDirectory + "DKPReport-bbcode.txt", FileMode.Create);
                sw2 = new StreamWriter(fs2);
            }
            catch (Exception ex)
            {
                MessageBox.Show("File IO Error. \n(" + ex.Message + ")", "Error");
            }

            // Events + Loot table
            sw.Write("<table width ='100%' border='1' cellspacing='0' cellpadding='3' bordercolor='#0000ff' bordercolorlight='#000000' bordercolordark='#ffffff' frame='border' rules='all' class='gensmall'>");
            sw.Write("<tr><td colspan='4' bgcolor='black'><b>Events and Loot</b></td></tr>");
            sw2.Write("[p][table][tr][th][b]Events and Loot[/b][/th][/tr]");
            foreach (DataRow ev in dtEL.Select("EventName NOT LIKE '%1'", "RaidNumber ASC"))
            {
                dkp = 0;
                twotier = false;
                ppl = dtDKS.Select("EventNameOrLoot='" + ev["EventName"] + "'").Length;
                string evstr = (string)ev["EventName"];
                if (Regex.IsMatch(evstr, ".*?0"))
                {
                    evstr = evstr.TrimEnd(new char[] { '0' });
                    twotier = true;
                }
                sw.Write("<tr><td colspan='4' bgcolor='black'><b>" + i + ".</b> " + evstr);
                sw2.Write("[tr][th][b]" + i + ".[/b] " + evstr);
                if (twotier)
                {
                    sw.Write("<b><i> Double Tier</b></i>");
                    sw2.Write("[b][i] Double Tier[/b][/i]");
                }
                sw.Write("</td></tr>");
                sw2.Write("[/th][/tr]");
                sw.Write("<tr><td width='30%'><i><b>Loot</b></i></td><td width='25%'><b><i>Recipient</b></i></td><td width='15%'><b><i>Cost</b></i></td><td width='30%'><b><i>DKP Per Person</b></i></td></tr>");
                sw2.Write("[tr][td][i][b]Loot[/b][/i][/td][td][b][i]Recipient[/b][/i][/td][td][b][i]Cost[/b][/i][/td][td][b][i]DKP Per Person[/b][/i][/td][/tr]");
                DataRow[] loots = dtDKS.Select("LootRaid='" + ev["EventName"] + "'");
                i++;
                foreach (DataRow r in loots)
                {
                    dkp += (double)r["PTS"] * -1;
                    sw.Write("<tr><td>" + r["EventNameOrLoot"] + "</td><td>" + r["Name"] + "</td><td>" + r["PTS"] + "</td><td>" + (double)r["PTS"] * -1 / ppl * (1 - owner.DKPTax) + "</td></tr>");
                    sw2.Write("[tr][td]" + r["EventNameOrLoot"] + "[/td][td]" + r["Name"] + "[/td][td]" + r["PTS"] + "[/td][td]" + (double)r["PTS"] * -1 / ppl * (1 - owner.DKPTax) + "[/td][/tr]");
                }
                sw.Write("<tr><td><b><i>Totals:</b></i></td><td>-</td><td>" + dkp * -1 + "</td><td>" + dkp / ppl * (1 - owner.DKPTax) + "</td></tr>");
                sw2.Write("[tr][td][b][i]Totals:[/b][/i][/td][td]-[/td][td]" + dkp * -1 + "[/td][td]" + dkp / ppl * (1 - owner.DKPTax) + "[/td][/tr]");
            }
            sw.Write("</table>\n\n");
            sw2.Write("[/table][/p]");
            owner.PBVal += 10;

            // Attendance table
            sw.Write("<table width ='100%' border='1' cellspacing='0' cellpadding='3' bordercolor='#0000ff' bordercolorlight='#000000' bordercolordark='#ffffff' frame='border' rules='all' class='gensmall'>");
            sw2.Write("[p][table]");
            i = 1;
            sw.Write("<tr><td colspan='" + (6 + dtEL.Select("EventName NOT LIKE '%1'").Length) + "' bgcolor='black'><b>Attendance</b></td></tr>");
            sw2.Write("[tr][th][b]Attendance[/b][/th][/tr]");
            sw.Write("<tr><td width='15%'><b><i>Name</b></i></td>");
            sw2.Write("[tr][td][b][i]Name[/b][/i][/td]");
            foreach (DataRow ev in dtEL.Select("EventName NOT LIKE '%1'"))
            {
                sw.Write("<td width='15'><b><i>" + i + "</b></i></td>");
                sw2.Write("[td][b][i]" + i + "[/i][/b][/td]");
                i++;
            }
            sw.Write("<td width='15%'><b><i>DKP Net Change</b></i></td><td width='15%'><b><i>DKP Total</b></i><td width='15'></td><td width='*'><b><i>Tier</b></i></td><td width='10'><i><b>MIA</i></b></td></tr>");
            sw2.Write("[td][b][i]DKP Net Change[/i][/b][/td][td][b][i]DKP Total[/i][/b][/td][td][/td][td][b][i]Tier[/i][/b][/td][td][b][i]MIA[/i][/b][/td][/tr]");
            owner.PBVal += 10;
            foreach (DataRow r in dtNT.Rows)
            {
                earned = 0;
                if (dtTotal.Select("Name='" + r["Name"] + "'").Length == 0) continue;

                sw.Write("<td>" + r["Name"] + "</td>");
                sw2.Write("[tr][td]" + r["Name"] + "[/td]");
                int event_counter = 1;
                foreach (DataRow ev in dtEL.Select("EventName NOT LIKE '%1'", "RaidNumber ASC"))
                {
                    if (dtDKS.Select("Name='" + r["Name"] + "' AND EventNameOrLoot='" + ev["EventName"] + "'").Length > 0)
                    {
                        earned += (double)dtDKS.Select("Name='" + r["Name"] + "' AND EventNameOrLoot='" + ev["EventName"] + "'")[0]["PTS"];
                        sw.Write("<td>•</td>");
                        sw2.Write("[td]" + event_counter + "[/td]");
                    }
                    else
                    {
                        sw.Write("<td></td>");
                        sw2.Write("[td][/td]");
                    }
                    event_counter++;
                }
                foreach (DataRow deduct in dtDKS.Select("Name='" + r["Name"] + "' AND PTS<0")) earned += (double)deduct["PTS"];
                totalnetchange += earned;
                totaldkp += (double)dtTotal.Select("Name='" + r["Name"] + "'")[0]["Total"];
                DateTime lastRaid = (DateTime)dtFinal.Select("Name='" + r["Name"] + "'")[0]["LastRaid"];
                int daysSinceLastRaid = (int)(t - lastRaid).TotalDays;

                sw.Write("<td>" + String.Format("{0:+0.00;-0.00;00.00}",earned) + "</td>");
				sw.Write("<td>" + String.Format("{0:+0.000#;-0.000#;00.000#}",dtTotal.Select("Name='" + r["Name"]+ "'")[0]["Total"]) + "</td>");
                sw.Write("<td>" + r["Tier"] + "</td><td>" + String.Format("{0:00.00}", r["TPercent"]) + "% (" + r["NumRaids"] + "/" + dtAllEvents.Rows.Count + ")</td>");
                sw.Write("<td>" + (daysSinceLastRaid >= owner.LastRaidDaysThreshold ? $"{daysSinceLastRaid}d" : "") + "</td></tr>");
                sw2.Write("[td]" + String.Format("{0:+0.00;-0.00;00.00}",earned) + "[/td]");
				sw2.Write("[td]" + String.Format("{0:+0.000#;-0.000#;00.000#}",dtTotal.Select("Name='" + r["Name"]+ "'")[0]["Total"]) + "[/td]");
                sw2.Write("[td]" + r["Tier"] + "[/td][td]" + String.Format("{0:00.00}", r["TPercent"]) + "% (" + r["NumRaids"] + "/" + dtAllEvents.Rows.Count + ")[/td]");
                sw2.Write("[td]" + (daysSinceLastRaid >= owner.LastRaidDaysThreshold ? $"{daysSinceLastRaid}d" : "") + "[/td][/tr]");
			}
			sw.Write("<tr><td><b><i>Totals:</b></i></td>");
            sw2.Write("[tr][td][b][i]Totals:[/i][/b][/td]");
			foreach (DataRow ev in dtEL.Select("EventName NOT LIKE '%1'")) 
            {
                sw.Write("<td></td>");
                sw2.Write("[td][/td]");
            }
			sw.Write("<td>" + String.Format("{0:+0.00;-0.00;00.00}",totalnetchange) + "</td>");
			sw.Write("<td>" + String.Format("{0:0.0#}",totaldkp) + "</td><td></td><td></td><td></td></tr>");
			sw.Write("</table>\n\n");
            sw2.Write("[td]" + String.Format("{0:+0.00;-0.00;00.00}",totalnetchange) + "[/td]");
			sw2.Write("[td]" + String.Format("{0:0.0#}",totaldkp) + "[/td][td][/td][td][/td][td][/td][/tr]");
			sw2.Write("[/table][/p]\n\n");
            owner.PBVal += 10;


            // Tier table
			DataRow[] a = dtFinal.Select("Tier='A'","Total DESC");
			DataRow[] b = dtFinal.Select("Tier='B'","Total DESC");
			DataRow[] c = dtFinal.Select("Tier='C'","Total DESC");
			int j =0;
                        
			sw.Write("<table width ='100%' border='1' cellspacing='0' cellpadding='3' bordercolor='#0000ff' bordercolorlight='#000000' bordercolordark='#ffffff' frame='border' rules='all' class='gensmall'>");
			sw.Write("<tr><td colspan='3' bgcolor='black'><b>Tier A</b></td><td colspan='3' bgcolor='black'><b>Tier B</b></td><td colspan='3' bgcolor='black'><b>Tier C</b></td></tr>");
			sw.Write("<tr><td width='15%'><b><i>Name</b></i></td><td width='5%'><b><i>%</b></i></td><td width='13%'><b><i>DKP Total</b></i></td><td width='15%'><b><i>Name</b></i></td><td width='5%'><b><i>%</b></i></td><td width='13%'><b><i>DKP Total</b></i></td><td width='15%'><b><i>Name</b></i></td><td width='5%'><b><i>%</b></i></td><td width='13%'><b><i>DKP Total</b></i></td></tr>");
            sw2.Write("[p][table][tr][th][b]Tier A[/b][/th][th][/th][th][/th][th][b]Tier B[/b][/th][th][/th][th][/th][th]>[b]Tier C[/b][/th][th][/th][th][/th][/tr]");
            sw2.Write("[tr][td][b][i]Name[/i][/b][/td][td][b][i]%[/i][/b][/td][td][b][i]DKP Total[/i][/b][/td][td][b][i]Name[/i][/b][/td][td][b][i]%[/i][/b][/td][td][b][i]DKP Total[/i][/b][/td][td][b][i]Name[/i][/b][/td][td][b][i]%[/i][/b][/td][td][b][i]DKP Total[/i][/b][/td][/tr]");
			double atd, btd, ctd, att, btt, ctt;
			atd = btd = ctd = att = btt = ctt = 0.00;

			while (a.Length>j||b.Length>j||c.Length>j)
			{
				sw.Write("<tr>");
                sw2.Write("[tr]");
                if (a.Length > j)
                {
                    sw.Write("<td>" + a[j]["Name"] + "</td><td>" + String.Format("{0:00.00}", a[j]["TPercent"]) + "</td><td>" + String.Format("{0:+0.0#;-0.0#;0.00}", a[j]["Total"]));
                    sw2.Write("[td]" + a[j]["Name"] + "[/td][td]" + String.Format("{0:00.00}", a[j]["TPercent"]) + "[/td][td]" + String.Format("{0:+0.0#;-0.0#;0.00}", a[j]["Total"]) + "[/td]");
                    atd += (double)a[j]["Total"];
                    att += double.Parse(a[j]["TPercent"].ToString());
                }
                else
                {
                    sw.Write("<td></td><td></td><td></td>");
                    sw2.Write("[td][/td][td][/td][td][/td]");
                }

                if (b.Length > j)
                {
                    sw.Write("<td>" + b[j]["Name"] + "</td><td>" + String.Format("{0:00.00}", b[j]["TPercent"]) + "</td><td>" + String.Format("{0:+0.0#;-0.0#;0.00}", b[j]["Total"]));
                    sw2.Write("[td]" + b[j]["Name"] + "[/td][td]" + String.Format("{0:00.00}", b[j]["TPercent"]) + "[/td][td]" + String.Format("{0:+0.0#;-0.0#;0.00}", b[j]["Total"]) + "[/td]");
                    btd += (double)b[j]["Total"];
                    btt += double.Parse(b[j]["TPercent"].ToString());
                }
                else
                {
                    sw.Write("<td></td><td></td><td></td>");
                    sw2.Write("[td][/td][td][/td][td][/td]");
                }

                if (c.Length > j)
                {
                    sw.Write("<td>" + c[j]["Name"] + "</td><td>" + String.Format("{0:00.00}", c[j]["TPercent"]) + "</td><td>" + String.Format("{0:+0.0#;-0.0#;0.00}", c[j]["Total"]));
                    sw2.Write("[td]" + c[j]["Name"] + "[/td][td]" + String.Format("{0:00.00}", c[j]["TPercent"]) + "[/td][td]" + String.Format("{0:+0.0#;-0.0#;0.00}", c[j]["Total"]) + "[/td]");
                    ctd += (double)c[j]["Total"];
                    ctt += double.Parse(c[j]["TPercent"].ToString());
                }
                else
                {
                    sw.Write("<td></td><td></td><td></td>");
                    sw2.Write("[td][/td][td][/td][td][/td]");
                }
				sw.Write("</tr>");
                sw2.Write("[/tr]");
				j++;
			}
			sw.Write("<tr><td><b><i>Averages:</i></b></td><td>" 
				+ String.Format("{0:00.00}",att/a.Length) + "</td><td>" + String.Format("{0:+0.0#;-0.0#;0.00}",atd/a.Length) + "</td><td></td><td>" 
				+ String.Format("{0:00.00}",btt/b.Length) + "</td><td>" + String.Format("{0:+0.0#;-0.0#;0.00}",btd/b.Length) + "</td><td></td><td>"  
				+ String.Format("{0:00.00}",ctt/c.Length) + "</td><td>" + String.Format("{0:+0.0#;-0.0#;0.00}",ctd/c.Length) + "</td></tr>");

            sw2.Write("[tr][td][b][i]Averages:[/i][/b][/td][td]"
                + String.Format("{0:00.00}", att / a.Length) + "[/td][td]" + String.Format("{0:+0.0#;-0.0#;0.00}", atd / a.Length) + "[/td][td][/td][td]"
                + String.Format("{0:00.00}", btt / b.Length) + "[/td][td]" + String.Format("{0:+0.0#;-0.0#;0.00}", btd / b.Length) + "[/td][td][/td][td]"
                + String.Format("{0:00.00}", ctt / c.Length) + "[/td][td]" + String.Format("{0:+0.0#;-0.0#;0.00}", ctd / c.Length) + "[/td][/tr]");

			sw.Write("</table>\n\n");
            sw2.Write("[/table][/p]\n\n");
            owner.PBVal += 10;

            // Events Table
			sw.Write("<table width ='100%' border='1' cellspacing='0' cellpadding='3' bordercolor='#0000ff' bordercolorlight='#000000' bordercolordark='#ffffff' frame='border' rules='all' class='gensmall'>");
			sw.Write("<tr><td colspan='2' bgcolor='black'><b>Events Counted Towards Tier</b></td>");
            sw2.Write("[p][table][tr][th][b]Events Counted Towards Tier[/b][/th][th][/th][/tr]");
            sw2.Write("[tr][td]");
			DateTime n = new DateTime(1,1,1);
  
			foreach (DataRow r in dtAllEvents.Rows)
			{
				if (!((DateTime)r["EventDate"]).Equals(n))
				{
					n = (DateTime)r["EventDate"];
					sw.Write("</td></tr><tr><td>" + n.Month + "/" + n.Day + "/" + n.Year + "</td><td>");
                    sw2.Write("[/td][/tr][tr][td]" + n.Month + "/" + n.Day + "/" + n.Year + "[/td][td]");
				}
				sw.Write(r["EventName"] + "\n");
                sw2.Write(r["EventName"] + "\n");
			}
			sw.Write("</td></tr></table>");
            sw2.Write("[/td][/tr][/table][/p]");
            owner.PBVal += 10;
                                                                                               
			sw.Close();
			fs.Close();
            sw2.Close();
            fs2.Close();
            owner.StatusMessage = "Finished writing daily report...";
        }
        #endregion
    }
}
