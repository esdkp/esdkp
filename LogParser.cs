using System;
using System.Windows.Forms;
using System.Collections;
using System.IO;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.OleDb;


namespace ES_DKP_Utils
{

   

	public class LogParser
    {
        public enum TriVal { TELL, WHO, NA }

        #region Declarations
        private FileStream logStream;
		private long logPointer;
		private long logSize;
		private System.IO.FileSystemWatcher logWatcher;
		private DataTable namesTiers;
		private string log;
        private bool changed;
        private System.Timers.Timer logTimer;

        private ArrayList _Tells;
		public ArrayList Tells
		{
			get
			{
				return _Tells;
			}
			set
			{
				_Tells = value;
                
			}
		}

		private ArrayList _TellsDKP;
		public ArrayList TellsDKP
		{
			get
			{
				return _TellsDKP;
			}
			set
			{
				_TellsDKP = value;
                
			}
		}

		private bool _TellsOn;
		public bool TellsOn
		{
			get
			{
				return _TellsOn;
			}
			set
			{
				_TellsOn = value;
                
			}
		}

		private bool _AttendanceOn;
		public bool AttendanceOn
		{
			get
			{
				return _AttendanceOn;
			}
			set
			{
				_AttendanceOn = value;
                
			}
		}

		private frmMain owner;
        private DebugLogger debugLogger;
        #endregion

        #region Constructor
        public LogParser(frmMain owner, string log)
		{
#if (DEBUG_1||DEBUG_2||DEBUG_3)
            debugLogger = new DebugLogger("LogParser.log");
#endif
            debugLogger.WriteDebug_3("Begin Method: LogParser.LogParser(frmMain,log) (" + owner.ToString() + "," + log.ToString() + ")");

			this.owner = owner;
			this.log = log;
			try 
			{
				logStream = new FileStream(log,FileMode.OpenOrCreate,FileAccess.Read);
				logPointer  = logSize = logStream.Length;
				logStream.Close();
			} 
			catch (Exception ex)
			{
                debugLogger.WriteDebug_1("Failed to initialize log parser: " + ex.Message);
				MessageBox.Show("Error initializing log parser.\n\n" + ex.Message,"Error");
			}
			try 
			{
				logWatcher = new FileSystemWatcher(Path.GetDirectoryName(log),Path.GetFileName(log));
				logWatcher.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.FileName | NotifyFilters.Size;
				logWatcher.Changed += new System.IO.FileSystemEventHandler(OnChanged);
				logWatcher.EnableRaisingEvents = true;
			}
			catch (Exception ex)
			{
                debugLogger.WriteDebug_1("Failed to initialize log watcher: " + ex.Message);
				MessageBox.Show("Error initializing file system watcher.\n\n" + ex.Message,"Error");
			}
			TellsOn = false;
			AttendanceOn = false;
			Tells = new ArrayList();
			TellsDKP = new ArrayList();
            logTimer = new System.Timers.Timer(5000);
            logTimer.Start();
            logTimer.Elapsed += new System.Timers.ElapsedEventHandler(logTimer_Elapsed);

			LoadNamesTiers();

            debugLogger.WriteDebug_3("End Method: LogParser.LogParser()");
        }
        #endregion

        #region Methods

        public void LoadNamesTiers()
		{
            debugLogger.WriteDebug_3("Begin Method: LogParser.LoadNamesTiers()");

			namesTiers = new DataTable();
			OleDbConnection dbConnect = null;
			try 
			{
				dbConnect = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + owner.DBString);
				string query = "Select NamesTiers.Name, NamesTiers.Tier, NamesTiers.TPercent, Sum(DKS.PTS) as SumOfPTS From (NamesTiers INNER JOIN DKS ON NamesTiers.Name = DKS.Name) GROUP BY NamesTiers.Name, NamesTiers.Tier, NamesTiers.TPercent";
				OleDbDataAdapter dkpDA = new OleDbDataAdapter(query,dbConnect);
				dbConnect.Open();
				dkpDA.Fill(namesTiers);
			}
			catch (Exception ex) 
			{
                debugLogger.WriteDebug_1("Failed to load NamesTiers table: " + ex.Message);
				MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")","Error");
			}
			finally { dbConnect.Close(); }

            debugLogger.WriteDebug_3("End Method: LogParser.LoadNamesTiers()");
		}

		private string getLine(byte[] data, int offset) 
		{
            debugLogger.WriteDebug_3("Begin Method: LogParser.getLine()");

            int i = 0;
			string s = "";
          
            while ( (i+offset)<data.Length && (char)data[i+offset] != '\n') 
            {
                s += (char)data[i+offset];
                i++;
            }

            owner.LineCount++;

            debugLogger.WriteDebug_3("End Method: LogParser.getLine(), returning " + s);
			return s;
		}

		public void OnChanged(object source, FileSystemEventArgs e)
		{
            debugLogger.WriteDebug_3("Begin Method: OnChanged(object,FileSystemEventArgs) (" 
                + source.ToString() + "," + e.ToString() + ")");

            changed = true;

            debugLogger.WriteDebug_3("End Method: OnChanged()");
		}

		public TriVal Parse(string s)
		{
            debugLogger.WriteDebug_3("Begin Method: Parse(string) (" + s.ToString() + ")");

			if (TellsOn) 
			{
				string name;
				Regex r = new Regex("\\[.*?\\]\\s[a-zA-Z]*?\\s(tells|told)\\syou,\\s'.*");
				Match m = r.Match(s);
				if (m.Success) 
				{
                    debugLogger.WriteDebug_3(s + " matches tell regex.");

                    owner.ParseCount++;
					name = Regex.Replace(s,"\\[.*?\\]\\s","");
					name = Regex.Replace(name,"\\s(tells|told)\\syou,\\s'.*","");
					if (!Tells.Contains(name)) 
					{
                        debugLogger.WriteDebug_3(name + " is not in tell array already.");

						Tells.Add(name);
						DataRow[] k = namesTiers.Select("Name='" + name + "'");
						if (k.Length>0) 
						{
							TellsDKP.Add(new Raider(owner,(string)k[0]["Name"],(string)k[0]["Tier"],(double)k[0]["SumOfPTS"],(double)Decimal.ToDouble((decimal)k[0]["TPercent"])));
						} 
						else 
						{
							TellsDKP.Add(new Raider(owner,name,Raider.NOTIER,Raider.NODKP,Raider.NOATTENDANCE));
						}
						TellsDKP.Sort();
                        owner.RefreshTells = true;					
					}
					return TriVal.TELL;
				}
                debugLogger.WriteDebug_3(s + " does not match tell regex.");
			}

			if (AttendanceOn)
			{
				string[] z = owner.Zones;
				Regex regex = new Regex("\\[.*?\\]\\s\\[.*?\\]\\s[a-zA-Z]*?\\s\\(.*?\\)\\s<(" + owner.GuildNames + ")>\\sZONE:\\s[a-z]*.*");
				Match m = regex.Match(s);
				if (m.Success)
				{
                    debugLogger.WriteDebug_3(s + " matches attendance regex.");

					string[] parsed = owner.CurrentRaid.ParseLine(s);
					foreach (string zone in z)
					{
						if (zone == parsed[1]) 
						{
                            debugLogger.WriteDebug_3(parsed[0] + " is in zone " + parsed[1] + " which is in attendance zone array, adding");
							owner.CurrentRaid.InputPerson(parsed[0]);
						} 
					}
                    owner.ParseCount++;
                    return TriVal.WHO;
				} else {
                    debugLogger.WriteDebug_3(s + " does not match attendance regex.");
                }
			}
            debugLogger.WriteDebug_3("End Method: Parse()");
            return TriVal.NA;
        }
        #endregion

        #region Events
        void logTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            debugLogger.WriteDebug_3("Begin Method: logTimer_Elapsed(object,ElapsedEventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

            if (!changed) return;
            changed = false;
            if (!(TellsOn || AttendanceOn)) return;
            byte[] newData = null;
            string line;
            int newSize = 0;
            int toRead = 0;

            FileStream logStream = null;

            logSize = new FileInfo(log).Length;
            try
            {
                logStream = new FileStream(log, FileMode.OpenOrCreate, FileAccess.Read);
            }
            catch (Exception ex)
            {
                debugLogger.WriteDebug_1("Failed to open log file: " + ex.Message);
            }
            try
            {
                newData = new byte[logStream.Length - logPointer + 1];
                logStream.Seek(logPointer, SeekOrigin.Begin);
                logStream.Read(newData, 0, (int)(logStream.Length - logPointer));
                toRead = newSize = (int)(logStream.Length - logPointer);
                logPointer = logStream.Length;
                owner.StatusMessage = "Log changes detected, parsing " + newSize + " bytes.";
                owner.PBMin = 0;
                owner.PBMax = newSize;
                owner.PBVal = 0;
            }
            catch (Exception ex)
            {
                debugLogger.WriteDebug_1("Failed to load new log data into memory: " + ex.Message);
            }
            finally
            {
                logStream.Close();
            }

            int tells = 0;
            int whos = 0;
            int lines = 0;

            while (toRead > 0)
            {
                line = getLine(newData, newSize - toRead);
                if (TellsOn || AttendanceOn)
                {
                    switch (Parse(line))
                    {
                        case TriVal.TELL:
                            tells++;
                            break;
                        case TriVal.WHO:
                            whos++;
                            break;
                        case TriVal.NA:
                            break;
                    }
                    lines++;
                }
                toRead -= line.Length+1;
                owner.PBVal += line.Length + 1;
            }

            owner.StatusMessage = "New lines: " + lines + ", " + tells + " tells, " + whos + " who results";

            debugLogger.WriteDebug_3("End Method: logTimer_Elapsed()");
        }
        #endregion
    }
}
