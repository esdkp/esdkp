using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Collections;

namespace ES_DKP_Utils
{
	public class Raid
    {

        #region Declarations
        private frmMain owner;

        private string OriginalRaidName;

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
        private DebugLogger debugLogger;
        #endregion

        #region Constuctor
        public Raid(frmMain owner)
		{
#if (DEBUG_1||DEBUG_2||DEBUG_3)
            debugLogger = new DebugLogger("Raid.log");
#endif
            debugLogger.WriteDebug_3("Begin Method: Raid.Raid(frmMain) (" + owner.ToString() + ")");

			this.owner = owner;
			this.ignore = new ArrayList();
			DoubleTier = false;
			string connectionStr;
			connectionStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + owner.DBString;
			try { dbConnect = new OleDbConnection(connectionStr); }
			catch (Exception ex) 
			{
                debugLogger.WriteDebug_1("Failed to create data connection: " + ex.Message);
				MessageBox.Show("Could not create data connection. \n(" + ex.Message + ")","Error");
			}
			dkpDA = new OleDbDataAdapter();
			dbRaid = new DataTable("Raid");
			dbRaid.Columns.Add("Name",System.Type.GetType("System.String"));
			dbRaid.Columns.Add("Date",System.Type.GetType("System.DateTime"));
			dbRaid.Columns.Add("EventNameOrLoot",System.Type.GetType("System.String"));
			dbRaid.Columns.Add("PTS",System.Type.GetType("System.Double"));
			dbRaid.Columns.Add("Comment",System.Type.GetType("System.String"));
			dbRaid.Columns.Add("LootRaid",System.Type.GetType("System.String"));

			dbPeople = new DataTable("People");
			try 
			{
				string query = "SELECT DISTINCT Name FROM DKS WHERE (Name<>\"zzzDKP Retired\")";
				dkpDA.SelectCommand = new OleDbCommand(query,dbConnect);
				dbConnect.Open();
				dkpDA.Fill(dbPeople);
			}
			catch (Exception ex) 
			{
                debugLogger.WriteDebug_1("Failed to load people table: " + ex.Message);
				MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")","Error");
			}
			finally { dbConnect.Close(); }

            debugLogger.WriteDebug_3("End Method: Raid.Raid()");
        }
        #endregion

        #region Methods
        public void LoadRaid()
		{
            debugLogger.WriteDebug_3("Begin Method: Raid.LoadRaid()");

			try
			{
				string query = "SELECT * FROM DKS WHERE (Date=#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "# AND (((EventNameOrLoot LIKE '" + RaidName + "%' AND LootRaid NOT LIKE '__%') OR LootRaid LIKE '" + RaidName + "%') AND EventNameOrLoot NOT LIKE '%1') )";
				dkpDA.SelectCommand = new OleDbCommand(query,dbConnect);
				dbConnect.Open();
				dkpDA.Fill(dbRaid);
				cmdBld = new OleDbCommandBuilder(dkpDA);
				Loaded = true;

				DataRow[] tiercheck =  dbRaid.Select("EventNameOrLoot LIKE '%0' OR LootRaid LIKE '%0'");
				if (tiercheck.Length>0) 
				{
					DoubleTier = true;
					foreach (DataRow r in tiercheck)
					{
						bool isloot = (double)r.ItemArray[3] < 0;
						if (isloot) 
						{
							string s = (string) r.ItemArray[5];
							s = s.TrimEnd(new char[] { '0' } );
							r.ItemArray = new object[] { 
														   r.ItemArray[0], r.ItemArray[1],
														   r.ItemArray[2], r.ItemArray[3],
														   r.ItemArray[4], s };
						}
						else 
						{
							string s = (string) r.ItemArray[2];
							s = s.TrimEnd(new char[] { '0' } );
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
                debugLogger.WriteDebug_1("Failed to load raid: " + ex.Message);
				MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")","Error");
			}
			finally { dbConnect.Close(); }

            debugLogger.WriteDebug_3("End Method: Raid.LoadRaid()");
		}

		public void ShowAttendees()
		{
            debugLogger.WriteDebug_3("Begin Method: Raid.ShowAttendees()");

			DataRow[] attendees = dbRaid.Select("(PTS>=0)");
			DataTable newTable = rowsToTable(dbRaid,attendees);
			TableViewer table = new TableViewer(owner,newTable,new TableViewer.UpdateDel(UpdateAttendees));
			table.Show();

            debugLogger.WriteDebug_3("End Method: Raid.ShowAttendees()");
		}
		public void UpdateAttendees(DataTable dt) 
		{
            debugLogger.WriteDebug_3("Begin Method: Raid.UpdateAttendees(DataTable) (" + dt.ToString() + ")");

			foreach (DataRow r in dbRaid.Select("(PTS>=0)")) { dbRaid.Rows.Remove(r); }
			foreach (DataRow r in dt.Rows) { dbRaid.Rows.Add(r.ItemArray); }			

            debugLogger.WriteDebug_3("End Method: Raid.UpdateAttendees()");
		}

		public void ShowLoot()
		{
            debugLogger.WriteDebug_3("Begin Method: Raid.ShowLoot()");

			DataRow[] loot = dbRaid.Select("(PTS<0)");
			DataTable newTable = rowsToTable(dbRaid,loot);
			TableViewer table = new TableViewer(owner,newTable,new TableViewer.UpdateDel(UpdateLoot));
			table.Show();

            debugLogger.WriteDebug_3("End Method: Raid.ShowLoot()");
		}

		public void UpdateLoot(DataTable dt) 
		{
            debugLogger.WriteDebug_3("Begin Method: Raid.UpdateLoot(DataTable) (" + dt.ToString() + ")");

			foreach (DataRow r in dbRaid.Select("(PTS<0)")) { dbRaid.Rows.Remove(r); }
			foreach (DataRow r in dt.Rows) { dbRaid.Rows.Add(r.ItemArray); }

            debugLogger.WriteDebug_3("End Method: Raid.UpdateLoot()");
		}

		public void InputPerson(string s)
		{
            debugLogger.WriteDebug_3("Begin Method: Raid.InputPerson(string) (" + s.ToString() + ")");

			DataRow[] rows = dbPeople.Select("Name='" + s + "'");

			if (rows.Length > 0)
			{
				DataRow[] rs = dbRaid.Select("Name='" + s + "' AND PTS>=0");
                if (rs.Length == 0)
                {
                    dbRaid.Rows.Add(new object[] { s, RaidDate, RaidName, 0, null, null });
                    debugLogger.WriteDebug_3("Added person: " + s);
                }
			}
			else
			{
				string t = GetAlt(s);
				if ((t != null && t.Length != 0)) 
				{ 
					if (dbRaid.Select("Name='" + t + "' AND PTS>=0").Length==0) 
					{
						DataRow[] rs = dbRaid.Select("Name='" + t + "' AND PTS>=0");
                        if (rs.Length == 0)
                        {
                            dbRaid.Rows.Add(new object[] { t, RaidDate, RaidName, 0, null, null });
                            debugLogger.WriteDebug_3("Added person: " + t);
                        }

					}
					
				}
				else 
				{
					if (!ignore.Contains(t)) 
					{
						AppAltDialog alt = new AppAltDialog(owner,s,dbPeople);
						t = alt.GetName();
						if ((t != null && t.Length != 0))
						{ 
							DataRow[] rs = dbRaid.Select("Name='" + t + "' AND PTS>=0");
							if (rs.Length==0) 
							{
								dbRaid.Rows.Add(new object[] {t,RaidDate,RaidName,0,null,null});
                                debugLogger.WriteDebug_3("Added person: " + t);
							}
							if (t!=s) dbAlts.Rows.Add(new object[] {s,t});
							altDA.Update(dbAlts);
						} 
						else { ignore.Add(t); }
					}
				}
			}
		}

		public void ParseAttendance(string logFile, string zo) 
		{
            debugLogger.WriteDebug_3("Begin Method: Raid.ParseAttendance(string,string) (" + logFile.ToString() + "," + zo.ToString() + ")");

			string[] zones = zo.Split(new char[] { ' ' });
			string line;
			FileStream input;
			StreamReader sr;
			Regex regex = new Regex("\\[.*?\\]\\s\\[.*?\\]\\s[a-zA-Z]*?\\s\\(.*?\\)\\s<Eternal Sovereign>\\sZONE:\\s[a-z]*.*");
			Match m;
			ArrayList people = new ArrayList();
		
			foreach (DataRow z in dbRaid.Rows)	people.Add(z.ItemArray[0]);

			try 
			{ 
				input = new FileStream(logFile,FileMode.Open,FileAccess.Read);
				sr = new StreamReader(input);
			}
			catch (Exception ex) 
			{
                debugLogger.WriteDebug_1("Failed to open log to parse: " + ex.Message);
				MessageBox.Show("File IO Error.\n(" + ex.Message + ")");
				return;
			}

			do
			{
				line = sr.ReadLine();
	
				if (line!=null) 
				{
					m = regex.Match(line);
                    if (m.Success)
                    {
                        debugLogger.WriteDebug_3(line + " matches attendance regex, parsing");
                        string[] parsed = ParseLine(line);
                        foreach (string zone in zones)
                        {
                            if ((zone == parsed[1]) && !people.Contains(parsed[0]))
                            {
                                debugLogger.WriteDebug_2(parsed[0] + " is in zone " + parsed[1] + ", adding to people array.");
                                people.Add(parsed[0]);
                            }
                        }
                    }
                    else
                    {
                        debugLogger.WriteDebug_3(line + " does not match attendance regex");
                    }
				}
			} while (line!=null);

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
                        debugLogger.WriteDebug_3("Added attendance row for " + s);
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
                                debugLogger.WriteDebug_3("Added attendance row for " + t + " (Alt: " + s + ")");
                            }
						}
					
					}
					else 
					{
						AppAltDialog alt = new AppAltDialog(owner,s,dbPeople);
						t = alt.GetName();
						if ((t != null && t.Length != 0))
						{ 
							DataRow[] rs = dbRaid.Select("Name='" + t + "' AND PTS>=0");
							if (rs.Length==0) 
							{
								dbRaid.Rows.Add(new object[] {t,RaidDate,RaidName,0,null,null});
                                debugLogger.WriteDebug_3("Added attendance row for " + t);
							}
							if (t!=s) dbAlts.Rows.Add(new object[] {s,t});
							altDA.Update(dbAlts);
						}
					}
				}
					
			}
            debugLogger.WriteDebug_3("Begin Method: Raid.ParseAttendance()");
		}

		public string GetAlt(string s)
		{
            debugLogger.WriteDebug_3("Begin Method: Raid.GetAlt(string) (" + s.ToString() + ")");

			dbAlts = new DataTable("Alts");
			string query = "SELECT * FROM ALTS";
			try 
			{
				altDA = new OleDbDataAdapter(query,dbConnect);
				dbConnect.Open();
				altDA.Fill(dbAlts);
				altBld = new OleDbCommandBuilder(altDA);
			}
			catch (Exception ex) 
			{
                debugLogger.WriteDebug_1("Failed to load alts table: " + ex.Message);
				MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")","Error");
			}
			finally { dbConnect.Close(); }
			
			DataRow[] rows = dbAlts.Select("AltName='" + s + "'");

            if (rows.Length > 0)
            {
                debugLogger.WriteDebug_3("End Method: Raid.GetAlt(), returning " + (string)rows[0].ItemArray[1]); ;
                return (string)rows[0].ItemArray[1];
            }

            debugLogger.WriteDebug_3("End Method: Raid.GetAlt(), returning ");
			return "";

		}

		public string[] ParseLine(string line)
		{
            debugLogger.WriteDebug_3("Begin Method: Raid.ParseLine(string) (" + line.ToString() + ")");

			string parsed = line.Trim();
			parsed = Regex.Replace(parsed,"\\[.*?\\]\\s\\[.*?\\]\\s","");
			parsed = Regex.Replace(parsed,"\\(.*?\\)\\s<Eternal Sovereign>\\sZONE:\\s","");
			string[] parsedSplit = parsed.Split(new char[] { ' ' });

            debugLogger.WriteDebug_3("End Method: Raid.ParseLine(), returning {" + parsedSplit[0] + "," + parsedSplit[1] + "}");
			return parsedSplit;
		}

		public void AddRowToLocalTable(object[] o)
		{
            debugLogger.WriteDebug_3("Begin Method: Raid.AddRowToLocalTable(object[]) (" + o.ToString() + ")");
            string s = "Creating new row: ";
            try
            {
                foreach (object k in o)
                {
                    if (k != null)
                        s += k.ToString() + ",";
                    else s += "null,";
                }
                debugLogger.WriteDebug_2(s.TrimEnd(new char[] { ',' }));
            }
            catch(Exception ex) { }
            

			DataRow r = dbRaid.NewRow();
			r.ItemArray = o;
			dbRaid.Rows.Add(r);

            debugLogger.WriteDebug_3("End Method: Raid.AddRowToLocalTable()");
		}

		public void DivideDKP()
		{
            debugLogger.WriteDebug_3("Begin Method: Raid.DivideDKP()");

			double dkp = 0.0;
			DataRow[] rows = dbRaid.Select("PTS<0 AND LootRaid='" + RaidName + "'");
			foreach (DataRow r in rows) { dkp += Math.Abs((double)r.ItemArray[3]); }
            debugLogger.WriteDebug_3("Total dkp to distribute: " + dkp);

			rows = dbRaid.Select("PTS>=0 AND EventNameOrLoot='" + RaidName + "'");
			if (rows.Length==0) rows = dbRaid.Select("PTS>=0 AND EventNameOrLoot='" + RaidName + "0'");
			if (rows.Length==0) return;

			Attendees = rows.Length;
            debugLogger.WriteDebug_3("Total attendees: " + Attendees);

			TotalDKP = dkp;
			dkp *= (1 - owner.DKPTax);
			dkp /= rows.Length;

            debugLogger.WriteDebug_3("DKP per person: " + dkp);

			foreach (DataRow r in rows) 
			{
				r.ItemArray = new object[] { r.ItemArray[0],r.ItemArray[1],r.ItemArray[2],dkp,r.ItemArray[4],r.ItemArray[5] };
			}

            debugLogger.WriteDebug_3("End Method: Raid.DivideDKP()");
		}

		public void SyncData() 
		{
            int k = 0;
            debugLogger.WriteDebug_3("Begin Method: Raid.SyncData()");

            owner.StatusMessage = "Saving raid...";
            owner.PBMin = 0;
            owner.PBMax = dbRaid.Rows.Count+1;
            owner.PBVal = 0;

			string query = "UPDATE DKS SET Comment='DeleteMe' WHERE Date=#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "# AND ( (EventNameOrLoot LIKE '" + RaidName + "%' AND PTS>=0)OR (LootRaid LIKE '" + RaidName + "%' AND PTS<0))";
			OleDbCommand updateCommand = null;
			OleDbCommand insertCommand  = null;
			OleDbCommand deleteCommand = null;
			try 
			{
				updateCommand = new OleDbCommand(query,dbConnect);
				dbConnect.Open();
				updateCommand.ExecuteNonQuery();
                debugLogger.WriteDebug_2("Updated " + k + " rows, flagged for deletion in DKS");

				query = "UPDATE EventLog SET DKPValue=-1 WHERE EventDate=#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "# AND ( EventName LIKE '" + RaidName + "%' )";
				updateCommand = new OleDbCommand(query,dbConnect);
				k = updateCommand.ExecuteNonQuery();
                debugLogger.WriteDebug_2("Updated " + k + " rows, flagged for deletion in EventLog");
			}
			catch (Exception ex) 
			{
                debugLogger.WriteDebug_1("Failed to load either DKS or EventLog: " + ex.Message);
				MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")","Error");
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
						updateCommand = new OleDbCommand(query,dbConnect);
						if ((k = updateCommand.ExecuteNonQuery()) == 0) 
						{
							query = "INSERT INTO DKS VALUES (\"" + r.ItemArray[0] + "\",#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "#,\"" + r.ItemArray[2] + "\"," + r.ItemArray[3] + ",\"" + r.ItemArray[4] + "\",\"" + r.ItemArray[5] + "\")";
							insertCommand = new OleDbCommand(query,dbConnect);
							k = insertCommand.ExecuteNonQuery();
						}
                        owner.PBVal++;
					}
					foreach (DataRow r in loot)
					{
						lootstr = (string)r.ItemArray[2];
						query = "INSERT INTO DKS VALUES (\"" + r.ItemArray[0] + "\",#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "#,\"" + lootstr + "\"," + r.ItemArray[3] + ",\"" + r.ItemArray[4] + "\",\"" + r.ItemArray[5] + "\")";
						insertCommand = new OleDbCommand(query,dbConnect);
						insertCommand.ExecuteNonQuery();

                        owner.PBVal++;
					}
					DataTable temp = new DataTable();
					query = "SELECT EventName FROM EventLog WHERE EventDate=#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "# AND EventName NOT LIKE '%1'";
					OleDbDataAdapter tempda =  new OleDbDataAdapter(query,dbConnect);
					tempda.Fill(temp);

					query = "UPDATE EventLog ";
					query += "SET Points=" + (TotalDKP/Attendees)*(1-owner.DKPTax) + ", DKPValue=" + TotalDKP + " ";
					query += "WHERE (EventDate=#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "# AND EventName=\"" + RaidName + "\")";

					updateCommand = new OleDbCommand(query,dbConnect);
					if (updateCommand.ExecuteNonQuery() == 0)
					{
						query = "INSERT INTO EventLog Values (#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "#,\"" + RaidName + "\"," + (TotalDKP/Attendees)*(1-owner.DKPTax) + "," + TotalDKP + "," + (int)((int)temp.Rows.Count + 1) + ")";
						insertCommand = new OleDbCommand(query,dbConnect);
						insertCommand.ExecuteNonQuery();
					}
                    owner.PBVal++;

					query = "DELETE FROM DKS WHERE Comment='DeleteMe'";
					deleteCommand = new OleDbCommand(query,dbConnect);
					deleteCommand.ExecuteNonQuery();
					query = "DELETE FROM EventLog WHERE DKPValue=-1";
					deleteCommand = new OleDbCommand(query,dbConnect);
					deleteCommand.ExecuteNonQuery();
				}
				catch (Exception ex) 
				{
                    debugLogger.WriteDebug_1("Failed to update DKS or EventLog: " + ex.Message);
					MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")","Error");
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
						updateCommand = new OleDbCommand(query,dbConnect);
						if (updateCommand.ExecuteNonQuery() == 0) 
						{
							query = "INSERT INTO DKS VALUES (\"" + r.ItemArray[0] + "\",#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "#,\"" + r.ItemArray[2] + "0\"," + r.ItemArray[3] + ",\"" + r.ItemArray[4] + "\",\"" + r.ItemArray[5] + "\")";
							insertCommand = new OleDbCommand(query,dbConnect);
							insertCommand.ExecuteNonQuery();
						}

						query = "UPDATE DKS SET EventNameOrLoot=\"" + r.ItemArray[2] + "1\", PTS=0, Comment=\"" + r.ItemArray[4] + "\", LootRaid=\"" + r.ItemArray[5] + "\" WHERE (Name=\"" + r.ItemArray[0] + "\" AND PTS>=0 AND Date=#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "# AND EventNameOrLoot=\"" + RaidName + "1\")";
						updateCommand = new OleDbCommand(query,dbConnect);
						if (updateCommand.ExecuteNonQuery() == 0) 
						{
							query = "INSERT INTO DKS VALUES (\"" + r.ItemArray[0] + "\",#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "#,\"" + r.ItemArray[2] + "1\",0,\"" + r.ItemArray[4] + "\",\"" + r.ItemArray[5] + "\")";
							insertCommand = new OleDbCommand(query,dbConnect);
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
						updateCommand = new OleDbCommand(query,dbConnect);

						if (updateCommand.ExecuteNonQuery() == 0) 
						{
							query = "INSERT INTO DKS VALUES (\"" + r.ItemArray[0] + "\",#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "#,\"" + lootstr + "\"," + r.ItemArray[3] + ",\"" + r.ItemArray[4] + "\",\"" + r.ItemArray[5] + "0\")";
							insertCommand = new OleDbCommand(query,dbConnect);
							insertCommand.ExecuteNonQuery();
						}

                        owner.PBVal++;
					}
					DataTable temp = new DataTable();
					query = "SELECT EventName FROM EventLog WHERE EventDate=#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "# AND EventName NOT LIKE '%1'";
					OleDbDataAdapter tempda =  new OleDbDataAdapter(query,dbConnect);
					tempda.Fill(temp);

					query = "UPDATE EventLog ";
					query += "SET Points=" + (TotalDKP/Attendees)*(1-owner.DKPTax) + ", DKPValue=" + TotalDKP + " ";
					query += "WHERE (EventDate=#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "# AND EventName=\"" + RaidName + "0\")";

					updateCommand = new OleDbCommand(query,dbConnect);
					if (updateCommand.ExecuteNonQuery() == 0)
					{
						query = "INSERT INTO EventLog Values (#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "#,\"" + RaidName + "0\"," + (TotalDKP/Attendees)*(1-owner.DKPTax) + "," + TotalDKP + "," + (int)((int)temp.Rows.Count + 1) + ")";
						insertCommand = new OleDbCommand(query,dbConnect);
						insertCommand.ExecuteNonQuery();
					}

					query = "UPDATE EventLog ";
					query += "SET Points=0, DKPValue=0 ";
					query += "WHERE (EventDate=#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "# AND EventName=\"" + RaidName + "1\")";

					updateCommand = new OleDbCommand(query,dbConnect);
					if (updateCommand.ExecuteNonQuery() == 0)
					{
						query = "INSERT INTO EventLog Values (#" + RaidDate.Month + "/" + RaidDate.Day + "/" + RaidDate.Year + "#,\"" + RaidName + "1\"," + 0 + "," + 0 + "," + (int)((int)temp.Rows.Count + 1) + ")";
						insertCommand = new OleDbCommand(query,dbConnect);
						insertCommand.ExecuteNonQuery();
					}

                    owner.PBVal++;

					query = "DELETE FROM DKS WHERE Comment='DeleteMe'";
					deleteCommand = new OleDbCommand(query,dbConnect);
					deleteCommand.ExecuteNonQuery();
					query = "DELETE FROM EventLog WHERE DKPValue=-1";
					deleteCommand = new OleDbCommand(query,dbConnect);
					deleteCommand.ExecuteNonQuery();
				}				
				catch (Exception ex) 
				{
                    debugLogger.WriteDebug_1("Failed to update DKS or EventLog: " + ex.Message);
					MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")","Error");
				}
				finally { dbConnect.Close(); }
			}

            debugLogger.WriteDebug_3("End Method: Raid.SyncRaid()");
		}

		public void FigureTiers()
		{
            debugLogger.WriteDebug_3("Begin Method: Raid.FigureTiers()");

			string query = "SELECT DISTINCT EventDate FROM EventLog WHERE EventDate <= #" + DateTime.Now.Month + "/" + DateTime.Now.Day + "/" + DateTime.Now.Year + "#";
			DataTable dates = null;
			DataTable nameCount = null;
			try
			{
				dates = new DataTable("Date");
				dkpDA = new OleDbDataAdapter(query,dbConnect);
				dbConnect.Open();
				dkpDA.Fill(dates);
			}
			catch (Exception ex) 
			{
                debugLogger.WriteDebug_1("Failed to fill dates table: " + ex.Message);
				MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")","Error");
			}
			finally { dbConnect.Close(); }

			DateTime tierwindow = (DateTime) dates.Rows[dates.Rows.Count - 20].ItemArray[0];
			int totalraids = 0;
			try
			{
				OleDbCommand deleteCmd = new OleDbCommand("DELETE * FROM NamesTiers",dbConnect);
				dbConnect.Open();
				deleteCmd.ExecuteNonQuery();
				nameCount = new DataTable("DKS");
				query = "SELECT DKS.Name, Count(DKS.Name) as NameCount FROM DKS WHERE DKS.Date>=#" + tierwindow.Month + "/" + tierwindow.Day + "/" + tierwindow.Year + "# AND DKS.PTS>=0 GROUP BY DKS.Name";
				dkpDA = new OleDbDataAdapter(query,dbConnect);
				dkpDA.Fill(nameCount);
				query = "SELECT Count(EventLog.EventName) as TotalRaids FROM EventLog WHERE EventLog.EventDate>=#" + tierwindow.Month + "/" + tierwindow.Day + "/" + tierwindow.Year + "#";
				dkpDA = new OleDbDataAdapter(query,dbConnect);
				DataTable t = new DataTable("temp");
				dkpDA.Fill(t);
				totalraids = (int)t.Rows[0].ItemArray[0];
			}
			catch (Exception ex) 
			{
                debugLogger.WriteDebug_1("Failed to fill attendance count or raid count table: " + ex.Message);
				MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")","Error");
			}
			finally { dbConnect.Close(); }

            owner.PBMin = 0;
            owner.PBMax = nameCount.Rows.Count + nameCount.Rows.Count / 10;
            owner.PBVal = nameCount.Rows.Count / 10;

			DataRow[] rows = nameCount.Select("Name<>'zzzDKP Retired'");
			foreach (DataRow r in rows)
			{
				int numraids = (int) r.ItemArray[1];
				char c;
				double d;
				string q;
				d = (double) numraids/totalraids;
				//if (d>=.60) c = 'A';
				//else if (d>.30) c = 'B';
				//else c = 'C';
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
					OleDbCommand dbc = new OleDbCommand(q,dbConnect);
					dbConnect.Open();
					dbc.ExecuteNonQuery();
                    owner.PBVal++;
				}
				catch (Exception ex) 
				{
                    debugLogger.WriteDebug_1("Failed to write NamesTiers row: " + ex.Message);
					MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")","Error");
				}
				finally { dbConnect.Close(); }

			}

            debugLogger.WriteDebug_3("End Method: Raid.FigureTiers()");
		}

		public void AddNameClass(string s, string t)
		{
            debugLogger.WriteDebug_3("Begin Method: Raid.AddNameClass(string,string) (" + s.ToString() + "," + t.ToString() + ")");

			string query = "INSERT INTO NamesClassesRemix VALUES ('" + s + "','" + t + "')";
			try
			{
				OleDbCommand nameclass = new OleDbCommand(query,dbConnect);
				dbConnect.Open();
				nameclass.ExecuteNonQuery();
			}
			catch (Exception ex) 
			{
                debugLogger.WriteDebug_1("Failed to insert NamesClasses row: " + ex.Message);
				MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")","Error");
			}
			finally { dbConnect.Close(); }

            debugLogger.WriteDebug_3("End Method: Raid.AddNameClass()");
		}

		public void AddPerson(string s)
		{
            debugLogger.WriteDebug_3("Begin Method: Raid.AddPerson(string) (" + s.ToString() + ")");

			if (Present().Select("Name='" + s + "' AND PTS>=0").Length==0) 
			{
				object[] o = new object[] { s, RaidDate, RaidName, 0, "", "" };
				AddRowToLocalTable(o);
				DivideDKP();
			}

            debugLogger.WriteDebug_3("End Method: Raid.AddPerson()");
		}

		public void RemovePerson(string s)
		{
            debugLogger.WriteDebug_3("Begin Method: Raid.RemovePerson(string) (" + s.ToString() + ")");

			DataRow[] rows = dbRaid.Select("Name='" + s + "' AND PTS>=0");
            try { dbRaid.Rows.Remove(rows[0]); }
            catch (Exception ex)
            {
                debugLogger.WriteDebug_1("Failed to remove " + s + " from local table: " + ex.Message);
            }
            DivideDKP();

            debugLogger.WriteDebug_3("End Method: Raid.RemovePerson()");
		}

		public DataTable NotPresent()
		{
            debugLogger.WriteDebug_3("Begin Method: Raid.NotPresent()");

			DataTable p = dbPeople.Copy();
			DataRow[] deleteMe = new DataRow[p.Rows.Count];
			int i = 0;
			int j = 0;
			foreach (DataRow r in p.Rows)
			{
				if (dbRaid.Select("Name='" + r.ItemArray[0] + "'").Length>0)
				{
					deleteMe[i] = r;
					i++;
				}
			}

			for (j=0; j<i; j++) p.Rows.Remove(deleteMe[j]);

            debugLogger.WriteDebug_3("End Method: Raid.NotPresent(), returning " + p.ToString());
			return p;
		}

		public DataTable Present()
		{
            debugLogger.WriteDebug_3("Begin Method: Raid.Present()");

			DataTable p = dbRaid.Copy();

            debugLogger.WriteDebug_3("End Method: Raid.Present(),  returning " + p);
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
			StreamWriter sw = null;
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
				MessageBox.Show("Could not create data connection. \n(" + ex.Message + ")","Error");
			}

			string query = "SELECT * FROM DKS WHERE Date=#" + t.Month + "/" + t.Day + "/" + t.Year + "# ORDER BY Name";
			try
			{
				OleDbDataAdapter da = new OleDbDataAdapter(query,dbCon);
				dbCon.Open();
				da.Fill(dtDKS);
				query = "SELECT * FROM EventLog WHERE EventDate=#" + t.Month + "/" + t.Day + "/" + t.Year + "# ORDER BY RaidNumber ASC";
				da = new OleDbDataAdapter(query,dbCon);
				da.Fill(dtEL);
                owner.PBVal += 10;
				query = "SELECT * FROM NamesTiers";
				da = new OleDbDataAdapter(query,dbCon);
				da.Fill(dtNT);
                owner.PBVal += 10;
				query = "SELECT Name, Sum(PTS) AS Total FROM DKS GROUP BY Name HAVING Name<>'zzzDKP Retired'";
				da = new OleDbDataAdapter(query,dbCon);
				da.Fill(dtTotal);
                owner.PBVal += 10;
				DataTable temp = new DataTable();
				query = "SELECT DISTINCT EventDate FROM EventLog ORDER BY EventDate DESC";
				da = new OleDbDataAdapter(query,dbCon);
				da.Fill(temp);
                owner.PBVal += 10;
				DateTime tempdate = (DateTime)temp.Rows[19]["EventDate"];
				query = "SELECT * FROM EventLog WHERE EventDate>=#" + tempdate.Month + "/" + tempdate.Day + "/" + tempdate.Year + "# ORDER BY EventDate ASC, RaidNumber ASC";
				da = new OleDbDataAdapter(query,dbCon);
				da.Fill(dtAllEvents);
                owner.PBVal += 10;
				query = "SELECT DKS.Name, Sum(DKS.PTS) AS Total, NamesTiers.Tier, NamesTiers.NumRaids, NamesTiers.TPercent ";
				query += "FROM DKS INNER JOIN NamesTiers ON DKS.Name=NamesTiers.Name GROUP BY DKS.Name, NamesTiers.Tier, NamesTiers.TPercent, NamesTiers.NumRaids";
				da = new OleDbDataAdapter(query,dbCon);
				da.Fill(dtFinal);
                owner.PBVal += 10;
			}
			catch (Exception ex) 
			{
				MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")","Error");
			}
			finally { dbCon.Close(); }

			try 
			{
				fs = new FileStream(owner.OutputDirectory + "DKPReport.txt",FileMode.Create);
				sw = new StreamWriter(fs);
			}
			catch (Exception ex) 
			{
				MessageBox.Show("File IO Error. \n(" + ex.Message + ")","Error");
			}

			sw.Write("<table width ='100%' border='1' cellspacing='0' cellpadding='3' bordercolor='#0000ff' bordercolorlight='#000000' bordercolordark='#ffffff' frame='border' rules='all' class='gensmall'>");
			sw.Write("<tr><td colspan='4' bgcolor='black'><b>Events and Loot</b></td></tr>");
			foreach (DataRow ev in dtEL.Select("EventName NOT LIKE '%1'","RaidNumber ASC"))
			{
				dkp=0;
				twotier=false;
				ppl = dtDKS.Select("EventNameOrLoot='" + ev["EventName"] + "'").Length;
				string evstr = (string)ev["EventName"];
				if (Regex.IsMatch(evstr,".*?0")) 
				{
					evstr = evstr.TrimEnd(new char[] { '0' });
					twotier = true;
				}
				sw.Write("<tr><td colspan='4' bgcolor='black'><b>" + i + ".</b> " + evstr);
				if (twotier) sw.Write("<b><i> Double Tier</b></i>");
				sw.Write("</td></tr>");
				sw.Write("<tr><td width='30%'><i><b>Loot</b></i></td><td width='25%'><b><i>Recipient</b></i></td><td width='15%'><b><i>Cost</b></i></td><td width='30%'><b><i>DKP Per Person</b></i></td></tr>");
				DataRow[] loots = dtDKS.Select("LootRaid='" + ev["EventName"] + "'");
				i++;
				foreach (DataRow r in loots) 
				{
					dkp += (double)r["PTS"] * -1;
					sw.Write("<tr><td>" + r["EventNameOrLoot"] + "</td><td>" + r["Name"] + "</td><td>" + r["PTS"] + "</td><td>" + (double)r["PTS"] * -1 / ppl * (1 - owner.DKPTax) + "</td></tr>");                  				
				}
				sw.Write("<tr><td><b><i>Totals:</b></i></td><td>-</td><td>" + dkp * -1 + "</td><td>" + dkp/ppl * (1 - owner.DKPTax) + "</td></tr>");
            }
			sw.Write("</table>\n\n");
            owner.PBVal += 10;
			sw.Write("<table width ='100%' border='1' cellspacing='0' cellpadding='3' bordercolor='#0000ff' bordercolorlight='#000000' bordercolordark='#ffffff' frame='border' rules='all' class='gensmall'>");
			i=1;
			sw.Write("<tr><td colspan='" + (5 + dtEL.Select("EventName NOT LIKE '%1'").Length) + "' bgcolor='black'><b>Attendance</b></td></tr>");
			sw.Write("<tr><td width='15%'><b><i>Name</b></i></td>");
			foreach (DataRow ev in dtEL.Select("EventName NOT LIKE '%1'")) sw.Write("<td width='15'><b><i>" + i++ + "</b></i>");
			sw.Write("<td width='15%'><b><i>DKP Net Change</b></i></td><td width='15%'><b><i>DKP Total</b></i><td width='15'></td><td width='*'><b><i>Tier</b></i></td></tr>");
            owner.PBVal += 10;
			foreach (DataRow r in dtNT.Rows)
			{
				earned = 0;
				if (dtTotal.Select("Name='" + r["Name"] + "'").Length==0) continue;

				sw.Write("<td>" + r["Name"] + "</td>");
				foreach (DataRow ev in dtEL.Select("EventName NOT LIKE '%1'","RaidNumber ASC"))
				{
					if (dtDKS.Select("Name='" + r["Name"] + "' AND EventNameOrLoot='" + ev["EventName"] + "'").Length>0) 
					{
						earned += (double)dtDKS.Select("Name='" + r["Name"] + "' AND EventNameOrLoot='" + ev["EventName"] + "'")[0]["PTS"];
						sw.Write("<td>•</td>");
					}
					else sw.Write("<td></td>");			
				}
				foreach (DataRow deduct in dtDKS.Select("Name='" + r["Name"] + "' AND PTS<0")) earned += (double)deduct["PTS"];
				totalnetchange += earned;
				totaldkp += (double)dtTotal.Select("Name='" + r["Name"] + "'")[0]["Total"];
				sw.Write("<td>" + String.Format("{0:+0.00;-0.00;00.00}",earned) + "</td>");
				sw.Write("<td>" + String.Format("{0:+0.000#;-0.000#;00.000#}",dtTotal.Select("Name='" + r["Name"]+ "'")[0]["Total"]) + "</td>");
				sw.Write("<td>" + r["Tier"] + "</td><td>" + String.Format("{0:00.00}",r["TPercent"]) + "% (" + r["NumRaids"] + "/" + dtAllEvents.Rows.Count + ")</td></tr>");
			}
			sw.Write("<tr><td><b><i>Totals:</b></i></td>");
			foreach (DataRow ev in dtEL.Select("EventName NOT LIKE '%1'")) sw.Write("<td></td>");
			sw.Write("<td>" + String.Format("{0:+0.00;-0.00;00.00}",totalnetchange) + "</td>");
			sw.Write("<td>" + String.Format("{0:0.0#}",totaldkp) + "</td><td></td><td></td></tr>");
			sw.Write("</table>\n\n");
            owner.PBVal += 10;

			DataRow[] a = dtFinal.Select("Tier='A'","Total DESC");
			DataRow[] b = dtFinal.Select("Tier='B'","Total DESC");
			DataRow[] c = dtFinal.Select("Tier='C'","Total DESC");
			int j =0;

			sw.Write("<table width ='100%' border='1' cellspacing='0' cellpadding='3' bordercolor='#0000ff' bordercolorlight='#000000' bordercolordark='#ffffff' frame='border' rules='all' class='gensmall'>");
			sw.Write("<tr><td colspan='3' bgcolor='black'><b>Tier A</b></td><td colspan='3' bgcolor='black'><b>Tier B</b></td><td colspan='3' bgcolor='black'><b>Tier C</b></td></tr>");
			sw.Write("<tr><td width='15%'><b><i>Name</b></i></td><td width='5%'><b><i>%</b></i></td><td width='13%'><b><i>DKP Total</b></i></td><td width='15%'><b><i>Name</b></i></td><td width='5%'><b><i>%</b></i></td><td width='13%'><b><i>DKP Total</b></i></td><td width='15%'><b><i>Name</b></i></td><td width='5%'><b><i>%</b></i></td><td width='13%'><b><i>DKP Total</b></i></td></tr>");
			double atd, btd, ctd, att, btt, ctt;
			atd = btd = ctd = att = btt = ctt = 0.00;

			while (a.Length>j||b.Length>j||c.Length>j)
			{
				sw.Write("<tr>");
				if (a.Length>j) 
				{
					sw.Write("<td>" + a[j]["Name"] + "</td><td>" + String.Format("{0:00.00}",a[j]["TPercent"]) + "</td><td>" + String.Format("{0:+0.0#;-0.0#;0.00}",a[j]["Total"]));
					atd += (double)a[j]["Total"];
					att += double.Parse(a[j]["TPercent"].ToString());
				}
				else sw.Write("<td></td><td></td><td></td>");

				if (b.Length>j) 
				{
					sw.Write("<td>" + b[j]["Name"] + "</td><td>" + String.Format("{0:00.00}",b[j]["TPercent"]) + "</td><td>" + String.Format("{0:+0.0#;-0.0#;0.00}",b[j]["Total"]));
					btd += (double)b[j]["Total"];
					btt += double.Parse(b[j]["TPercent"].ToString());
				}
				else sw.Write("<td></td><td></td><td></td>");

				if (c.Length>j) 
				{
					sw.Write("<td>" + c[j]["Name"] + "</td><td>" + String.Format("{0:00.00}",c[j]["TPercent"]) + "</td><td>" + String.Format("{0:+0.0#;-0.0#;0.00}",c[j]["Total"]));
					ctd += (double)c[j]["Total"];
					ctt += double.Parse(c[j]["TPercent"].ToString());
				}
				else sw.Write("<td></td><td></td><td></td>");
				sw.Write("</tr>");
				j++;
			}
			sw.Write("<tr><td><b><i>Averages:</i></b></td><td>" 
				+ String.Format("{0:00.00}",att/a.Length) + "</td><td>" + String.Format("{0:+0.0#;-0.0#;0.00}",atd/a.Length) + "</td><td></td><td>" 
				+ String.Format("{0:00.00}",btt/b.Length) + "</td><td>" + String.Format("{0:+0.0#;-0.0#;0.00}",btd/b.Length) + "</td><td></td><td>"  
				+ String.Format("{0:00.00}",ctt/c.Length) + "</td><td>" + String.Format("{0:+0.0#;-0.0#;0.00}",ctd/c.Length) + "</td></tr>");


			sw.Write("</table>\n\n");
            owner.PBVal += 10;
			sw.Write("<table width ='100%' border='1' cellspacing='0' cellpadding='3' bordercolor='#0000ff' bordercolorlight='#000000' bordercolordark='#ffffff' frame='border' rules='all' class='gensmall'>");
			sw.Write("<tr><td colspan='2' bgcolor='black'><b>Events Counted Towards Tier</b></td>");
			DateTime n = new DateTime(1,1,1);
  
			foreach (DataRow r in dtAllEvents.Rows)
			{
				if (!((DateTime)r["EventDate"]).Equals(n))
				{
					n = (DateTime)r["EventDate"];
					sw.Write("</td></tr><tr><td>" + n.Month + "/" + n.Day + "/" + n.Year + "</td><td>");
				}
				sw.Write(r["EventName"] + "\n");
			}
			sw.Write("</td></tr></table>");
            owner.PBVal += 10;
                                                                                               
			sw.Close();
			fs.Close();
            owner.StatusMessage = "Finished writing daily report...";
        }
        #endregion
    }
}
