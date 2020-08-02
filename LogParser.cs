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
        public enum LogLineType { TELL, WHO, LOOT, NA }

        #region Declarations
        private FileStream logStream;
        private long logPointer;
        private long logSize;
        private System.IO.FileSystemWatcher logWatcher;
        private DataTable namesTiers;
        private DataTable alts;
        private string log;
        private bool changed;
        private System.Timers.Timer logTimer;

        private ArrayList _Tells;
        public ArrayList Tells
        {
            get	{ return _Tells; }
            set	{ _Tells = value; }
        }

        private ArrayList _TellsDKP;
        public ArrayList TellsDKP
        {
            get { return _TellsDKP;	}
            set	{ _TellsDKP = value; }
        }

        private bool _TellsOn;
        public bool TellsOn
        {
            get	{ return _TellsOn; }
            set	{
                _TellsOn = value;
            }
        }

        private bool _AttendanceOn;
        public bool AttendanceOn
        {
            get { return _AttendanceOn; }
            set { _AttendanceOn = value; }
        }

        private bool _LootOn;
        public bool LootOn
        {
            get { return _LootOn; }
            set { _LootOn = value; }
        }

        private frmMain owner;
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        #endregion

        #region Constructor
        public LogParser(frmMain owner, string log)
        {
            logger.Debug("Begin Method: LogParser.LogParser(frmMain,log) (" + owner.ToString() + "," + log.ToString() + ")");

            this.owner = owner;
            this.log = log;

            // TODO: Was going to replace complicated logic with something like this, but will take time to test.
            //try
            //{
            //    using (FileStream fs = new FileStream(log, FileMode.OpenOrCreate, FileAccess.Read))
            //    {
            //        using (StreamReader sr = new StreamReader(fs)) {
            //            while (changed)
            //            {
            //                while (!sr.EndOfStream)
            //                    Parse(sr.ReadLine())
            //                while (sr.EndOfStream)
            //                    Thread.Sleep(100)
            //            }
            //        }
            //    }
            //}

            try 
            {
                logStream = new FileStream(log,FileMode.OpenOrCreate,FileAccess.Read);
                logPointer  = logSize = logStream.Length;
                logStream.Close();
            } 
            catch (Exception ex)
            {
                logger.Error("Failed to initialize log parser: " + ex.Message);
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
                logger.Error("Failed to initialize log watcher: " + ex.Message);
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

            logger.Debug("End Method: LogParser.LogParser()");
        }
        #endregion

        #region Methods

        public void LoadNamesTiers()
        {
            logger.Debug("Begin Method: LogParser.LoadNamesTiers()");

            OleDbConnection dbConnect = null;

            namesTiers = new DataTable();
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
                logger.Error("Failed to load NamesTiers table: " + ex.Message);
                MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")","Error");
            }
            finally { dbConnect.Close(); }

            alts = new DataTable();
            try
            {
                dbConnect = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + owner.DBString);
                string query = "SELECT AltName, Field3 FROM Alts";
                OleDbDataAdapter altDA = new OleDbDataAdapter(query, dbConnect);
                dbConnect.Open();
                altDA.Fill(alts);
            }
            catch (Exception ex)
            {
                logger.Error("Failed to load Alts table: " + ex.Message);
                MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")", "Error");
            }
            finally { dbConnect.Close(); }

            logger.Debug("End Method: LogParser.LoadNamesTiers()");
        }

        private string getLine(byte[] data, int offset) 
        {
            //logger.Debug("Begin Method: LogParser.getLine()");

            int i = 0;
            string s = "";
          
            while ( (i+offset)<data.Length && (char)data[i+offset] != '\n') 
            {
                s += (char)data[i+offset];
                i++;
            }

            owner.LineCount++;

            //logger.Debug("End Method: LogParser.getLine(), returning " + s);
            return s;
        }

        public void OnChanged(object source, FileSystemEventArgs e)
        {
            logger.Debug("Begin Method: OnChanged(object,FileSystemEventArgs) (" 
                + source.ToString() + "," + e.ToString() + ")");

            changed = true;

            logger.Debug("End Method: OnChanged()");
        }

        public LogLineType Parse(string s)
        {
            if (TellsOn)
            {
                Regex r = new Regex(@"\[.*\] (?<name>\S+) (tells|told) you, '(?<message>.*)'");
                Match m = r.Match(s);
                string tellType = "N/A";
                string[] keywords = { "main", "alt", "guest", "app" };

                if (m.Success) 
                {
                    logger.Debug(s + " matches tell regex.");

                    owner.ParseCount++;

                    string name = m.Groups["name"].ToString();
                    string message = m.Groups["message"].ToString();

                    foreach (string keyword in keywords)
                    {
                        if (message.StartsWith(keyword))
                        {
                            tellType = keyword;
                            message = message.Remove(0, keyword.Length);
                            message = message.TrimStart();
                        }
                    }

                    if (!Tells.Contains(name)) 
                    {
                        logger.Debug(name + " is not in tell array already.");

                        Tells.Add(name);
                        DataRow[] k = namesTiers.Select("Name='" + name + "'");
                        DataRow[] alt = alts.Select("AltName='" + name + "'");
                        
                        // If it's an alt, and no main was found, set the main array to the alt.
                        if (alt.Length > 0 && k.Length<=0)
                        {
                            k = namesTiers.Select("Name='" + (string)alt[0]["Field3"] + "'");
                        }

                        // Now if we successfully found someone in the namesTiers table, they should be a main
                        if (k.Length>0) 
                        {
                            TellsDKP.Add(new Raider(owner,(string)k[0]["Name"],(string)k[0]["Tier"],(double)k[0]["SumOfPTS"],(double)Decimal.ToDouble((decimal)k[0]["TPercent"]), tellType));
                        } 
                        // Otherwise, add them to the tells table with the default values for not tracked
                        else
                        {
                            TellsDKP.Add(new Raider(owner,name,Raider.NOTIER,Raider.NODKP,Raider.NOATTENDANCE, tellType));
                        }

                        TellsDKP.Sort();
                        owner.RefreshTells = true;					
                    }
                    return LogLineType.TELL;
                }
            }

            if (AttendanceOn)
            {
                string zones = owner.Zones;
                Regex r = new Regex(@"\[.*\] \[.*\] (?<name>\S+).*<(?<guild>.*)> ZONE: (?<zone>.*)$");
                Match m = r.Match(s.Trim());

                if (m.Success)
                {
                    string _guild = m.Groups["guild"].ToString();
                    string _zone = m.Groups["zone"].ToString();
                    string _name = m.Groups["name"].ToString();

                    if ( zones.Contains(_zone) && owner.GuildNames.Contains(_guild))
                    {
                        logger.Debug(_name + " <" + _guild + "> is in " + _zone + " which is in attendance zone array, adding");
                        owner.CurrentRaid.InputPerson(_name);
                    } 

                    owner.ParseCount++;
                    return LogLineType.WHO;
                }
            }
            return LogLineType.NA;
        }
        #endregion

        #region Events
        void logTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            logger.Debug("Begin Method: logTimer_Elapsed(object,ElapsedEventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

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
                logger.Error("Failed to open log file: " + ex.Message);
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
                logger.Error("Failed to load new log data into memory: " + ex.Message);
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
                        case LogLineType.TELL:
                            tells++;
                            break;
                        case LogLineType.WHO:
                            whos++;
                            break;
                        case LogLineType.NA:
                            break;
                    }
                    lines++;
                }
                toRead -= line.Length+1;
                owner.PBVal += line.Length + 1;
            }

            owner.StatusMessage = "New lines: " + lines + ", " + tells + " tells, " + whos + " who results";

            logger.Debug("End Method: logTimer_Elapsed()");
        }
        #endregion
    }
}
