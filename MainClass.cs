using Newtonsoft.Json;
using Nini.Config;
using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace ES_DKP_Utils
{
	/// <summary>
	/// DKP Utilities
	/// </summary>
	public class frmMain : System.Windows.Forms.Form
	{
        #region Menu Component Declarations
		private System.Windows.Forms.MenuItem mnuDKP;
		private System.Windows.Forms.MenuItem mnuSettings;
		private System.Windows.Forms.MenuItem mnuNewRaid;
		private System.Windows.Forms.MenuItem mnuOpenRaid;
		private System.Windows.Forms.MenuItem mnuReports;
		private System.Windows.Forms.MenuItem mnuDailyReport;
		private System.Windows.Forms.MenuItem mnuDetailedTier;
		private System.Windows.Forms.MenuItem mnuClassReport;
		private System.Windows.Forms.MenuItem mnuQueries;
		private System.Windows.Forms.MenuItem mnuBalanceQuery;
		private System.Windows.Forms.MenuItem mnuRetire;
		private System.Windows.Forms.MenuItem mnuUnretire;
		private System.Windows.Forms.MenuItem mnuEnterSQL;
		private System.Windows.Forms.MenuItem mnuAbout;
		private System.Windows.Forms.MenuItem mnuDKS;
		private System.Windows.Forms.MenuItem mnuAlts;
		private System.Windows.Forms.MenuItem mnuEventLog;
		private System.Windows.Forms.MenuItem mnuNamesTiers;
		private System.Windows.Forms.MenuItem mnuDKPTotals;
		private System.Windows.Forms.MenuItem mnuExit;
		private System.Windows.Forms.MenuItem mnuNameClass;
        private System.Windows.Forms.MainMenu mmuMain;
        private System.Windows.Forms.MenuItem mnuFileSep;
        private System.Windows.Forms.MenuItem mnuDKPOT;
        private System.Windows.Forms.MenuItem mnuTables;
        #endregion

        #region Other Component Declarations
        private System.Windows.Forms.StatusBar stbStatusBar;
		private System.Windows.Forms.StatusBarPanel sbpMessage;
		private System.Windows.Forms.StatusBarPanel sbpParseCount;
		private System.Windows.Forms.StatusBarPanel sbpLineCount;
        private System.Windows.Forms.Panel panel;
		private System.Windows.Forms.Label lblRaidName;
		private System.Windows.Forms.Label lblRaidDate;
		private System.Windows.Forms.TextBox txtRaidName;
		private System.Windows.Forms.Button btnAttendees;
		private System.Windows.Forms.Button btnLoot;
		private System.Windows.Forms.Label dkpLabel;
		private System.Windows.Forms.Label lblTier;
        private System.Windows.Forms.Label lblAttd;
		private System.Windows.Forms.ListBox listOfDKP;
		private System.Windows.Forms.ListBox listOfTiers;
		private System.Windows.Forms.ListBox listOfNames;
        private System.Windows.Forms.ListBox listOfAttd;
		private System.Windows.Forms.GroupBox grpItemTier;
		private System.Windows.Forms.RadioButton rdoA;
		private System.Windows.Forms.RadioButton rdoB;
		private System.Windows.Forms.RadioButton rdoC;
        private System.Windows.Forms.RadioButton rdoATTD;
		private System.Windows.Forms.Button btnRecordLoot;
		private System.Windows.Forms.Button btnSaveRaid;
		private System.Windows.Forms.TextBox txtZoneNames;
        private System.Windows.Forms.Label lblZoneNames;
		private System.Windows.Forms.CheckBox chkDouble;
		private System.Windows.Forms.Button btnAdd;
		private System.Windows.Forms.Button btnRemove;
        private System.Windows.Forms.DateTimePicker dtpRaidDate;
		private System.Windows.Forms.Label lblName;
        private System.Windows.Forms.StatusBarPanel sbpProgressBar;
        private System.Windows.Forms.ProgressBar pgbProgress;
        private System.Windows.Forms.Button btnClear;
        #endregion

        #region Properties
        public string Zones 
		{
			get 
			{
                return this.txtZoneNames.Text;
			}
		}

        private bool _RefreshTells;
        public bool RefreshTells
        {
            get { return _RefreshTells; }
            set { _RefreshTells = value; }
        }

        private string _LogFile;
		public string LogFile 
		{
			get { return _LogFile; }
			set 
			{ 
				try 
				{
					parser = new LogParser(this,value);
				}
				catch (Exception ex) 
				{
                    log.Debug("Could not create new log parser: " + ex.Message);
					MessageBox.Show("Error creating new log parser.  Log file not changed.", "Error");
				}
				_LogFile = value;
			}
		}

		private string _DBString;
		public string DBString
		{
			get	{ return _DBString; }
			set
			{
				_DBString = value;
                if (parser != null) parser.LoadNamesTiers();
			}
		}

		private string _OutputDirectory;
		public string OutputDirectory
		{
			get { return _OutputDirectory; }
			set { _OutputDirectory = value; }
		}

        private string _BackupDirectory;
        public string BackupDirectory
        {
            get { return _BackupDirectory; }
            set { _BackupDirectory = value; }
        }

        private System.Double _DKPTax;
		public System.Double DKPTax
		{
			get { return _DKPTax; }
			set { _DKPTax = value; }
		}

        private System.Double _MinDKP;
        /// <summary>
        /// Minimum balance to have tier considered (Tier 'D' limit, colloquially)
        /// </summary>
        public System.Double MinDKP
        {
            get { return _MinDKP; }
            set { _MinDKP = value; }
        }

        private System.Double _TierAPct;
        public System.Double TierAPct
        {
            get { return _TierAPct; }
            set { _TierAPct = value; }
        }

        private System.Double _TierBPct;
        public System.Double TierBPct
        {
            get { return _TierBPct; }
            set { _TierBPct = value; }
        }

        private System.Double _TierCPct;
        public System.Double TierCPct
        {
            get { return _TierCPct; }
            set { _TierCPct = value; }
        }

        private System.String _GuildNames;
        public System.String GuildNames
        {
            get { return _GuildNames; }
            set { _GuildNames = value; }
        }

        private System.Int32 _RaidDaysWindow;
        public System.Int32 RaidDaysWindow
        {
            get { return _RaidDaysWindow; }
            set { _RaidDaysWindow = value; }
        }

		private string _ItemDKP = "B";
		public string ItemDKP
		{
			get	{ return _ItemDKP; }
			set	{ _ItemDKP = value; }
		}

		private Raid _CurrentRaid;
		public Raid CurrentRaid 
		{
			get { return _CurrentRaid; }
            private set { _CurrentRaid = value; }
		}

        private string _StatusMessage;
        public string StatusMessage
        {
            get { return _StatusMessage; }
            set { _StatusMessage = value; }
        }

        private int _LineCount;
        public int LineCount
        {
            get { return _LineCount; }
            set { _LineCount = value; }
        }

        private int _ParseCount;
        public int ParseCount
        {
            get { return _ParseCount; }
            set { _ParseCount = value; }
        }

        private int _PBMin;
        public int PBMin
        {
            get { return _PBMin; }
            set { _PBMin = value; }
        }

        private int _PBMax;
        public int PBMax
        {
            get { return _PBMax; }
            set { _PBMax = value; }
        }

        private int _PBVal;
        public int PBVal
        {
            get { return _PBVal; }
            set { _PBVal = value; }
        }

        private bool _AutomaticBackups;
        public bool AutomaticBackups
        {
            get { return _AutomaticBackups; }
            set { _AutomaticBackups = value; }
        }
        #endregion

        public int LastRaidDaysThreshold { get; set; }
        public bool AutosaveJsonRaidModels { get; set; } = true;
        public string JsonRaidModelDirectory { get; set; } = Path.Combine(Directory.GetCurrentDirectory(), "jsonRaidModels");

        private IConfigSource inifile;
        private LogParser parser;
        private IContainer components;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private MenuItem mnuImport;
        private MenuItem mnuImportRaidDump;
        private MenuItem mnuImportLog;
        private GroupBox grpWatchFor;
        private CheckBox chkTells;
        private CheckBox chkWho;
        private CheckBox chkLoot;
        private MenuItem mnuExitNoBackup;
        private ListBox listTellType;
        private Label lblTellType;
        private ListBox listItemMessage;
        private Label lblMessage;
        private CheckBox chkSortIncludesItemMessage;
        private MenuItem mnuExport;
        private MenuItem mnuExportRaidJson;
        private MenuItem mnuExportAllRaidJson;
        private System.Windows.Forms.Timer UITimer;

        #region Constructor
        public frmMain()
		{
            log.Debug("Begin Method: frmMain.frmMain()");

			InitializeComponent();

			Directory.SetCurrentDirectory(Directory.GetParent(Application.ExecutablePath).FullName);
			log.Info("Set Current Directory to " + Directory.GetParent(Application.ExecutablePath).FullName);

            // TODO: Determine why this is a double nested try/catch to load some INI settings
			try 
			{
				inifile = new IniConfigSource(Directory.GetCurrentDirectory() + "\\settings.ini");

				DBString = inifile.Configs["Files"].GetString("dbfile", Directory.GetCurrentDirectory() + "\\DKP.mdb");
				LogFile = inifile.Configs["Files"].GetString("logfile", "");
				OutputDirectory = inifile.Configs["Files"].GetString("outdir", Directory.GetCurrentDirectory());
                BackupDirectory = inifile.Configs["Files"].GetString("backupdir", Directory.GetCurrentDirectory());
                JsonRaidModelDirectory = inifile.Configs["Files"].GetString("JsonRaidModelDirectory", JsonRaidModelDirectory);

                AutomaticBackups = inifile.Configs["Other"].GetBoolean("AutomaticBackups", true);
                AutosaveJsonRaidModels = inifile.Configs["Other"].GetBoolean("AutosaveJsonRaidModels", AutosaveJsonRaidModels);
                DKPTax = inifile.Configs["Other"].GetDouble("tax", 0.0);
                GuildNames = inifile.Configs["Other"].GetString("GuildNames", "Eternal Sovereign");
                LastRaidDaysThreshold = inifile.Configs["Other"].GetInt("LastRaidDaysThreshold", 7);
                MinDKP = inifile.Configs["Other"].GetDouble("mindkp", 0);
                RaidDaysWindow = inifile.Configs["Other"].GetInt("RaidDaysWindow", 20);
                TierAPct = inifile.Configs["Other"].GetDouble("tierapct", 0.60);
                TierBPct = inifile.Configs["Other"].GetDouble("tierbpct", 0.30);
                TierCPct = inifile.Configs["Other"].GetDouble("tiercpct", 0.01);

                log.Info("Read settings from INI: DBFile=" + DBString + ", LogFile=" + LogFile + ", OutputDirectory="
                    + OutputDirectory + ", DKPTax=" + DKPTax + ", GuildNames=" + GuildNames);

				StatusMessage ="Read settings from INI...";
				try 
				{
					parser = new LogParser(this,LogFile);
					log.Debug("Created log parser");
				} 
				catch (Exception ex) {
					log.Error("Failed to create log parser: " + ex.Message);
				}
			} 
			catch (FileNotFoundException ex) 
			{
                log.Debug(ex.Message);
				log.Info("settings.ini not found, creating");

				FileStream a = File.Create(Directory.GetCurrentDirectory() + "\\settings.ini");
				a.Close();
				sbpMessage.Text = "settings.ini not found... Loading defaults";
				try 
				{ 
					inifile = new IniConfigSource(Directory.GetCurrentDirectory() + "\\settings.ini");
					inifile.AddConfig("Files");
					inifile.AddConfig("Other");

					LogFile = inifile.Configs["Files"].GetString("logfile","");
					DBString = inifile.Configs["Files"].GetString("dbfile", Directory.GetCurrentDirectory() + "\\DKP.mdb");
					OutputDirectory = inifile.Configs["Files"].GetString("outdir", Directory.GetCurrentDirectory());
                    BackupDirectory = inifile.Configs["Files"].GetString("backupdir", Directory.GetCurrentDirectory());
                    JsonRaidModelDirectory = inifile.Configs["Files"].GetString("JsonRaidModelDirectory", JsonRaidModelDirectory);

                    AutomaticBackups = inifile.Configs["Other"].GetBoolean("AutomaticBackups", true);
                    AutosaveJsonRaidModels = inifile.Configs["Other"].GetBoolean("AutosaveJsonRaidModels", AutosaveJsonRaidModels);
                    DKPTax = inifile.Configs["Other"].GetDouble("tax",0.0);
                    GuildNames = inifile.Configs["Other"].GetString("GuildNames", "Eternal Sovereign");
                    LastRaidDaysThreshold = inifile.Configs["Other"].GetInt("LastRaidDaysThreshold", 7);
                    MinDKP = inifile.Configs["Other"].GetDouble("mindkp", 0);
                    RaidDaysWindow = inifile.Configs["Other"].GetInt("RaidDaysWindow", 20);
                    TierAPct = inifile.Configs["Other"].GetDouble("tierapct", 0.60);
                    TierBPct = inifile.Configs["Other"].GetDouble("tierbpct", 0.30);
                    TierCPct = inifile.Configs["Other"].GetDouble("tiercpct", 0.01);

                    inifile.Save();
					log.Info("Read settings from INI: dbFile=" + DBString + ", logFile=" + LogFile
                        + ", outDir=" + OutputDirectory + ", tax=" + DKPTax + ", mindkp=" + MinDKP + ", GuildNames=" + GuildNames);
				} 
				catch (Exception exc) 
				{ 
					log.Error("Failed to create new settings.ini: " + exc.Message);
					MessageBox.Show("Error opening settings.ini","Error");
					Environment.Exit(1);
				}
			}	
			catch (Exception ex) 
			{
				log.Error("Failed to load settings.ini: " + ex.Message);
				MessageBox.Show("Error opening settings.ini","Error");
				Environment.Exit(1);
			}

            if (Directory.Exists(BackupDirectory))
            {
                // TODO: add updating dkp database with newest version in backup directory
                //MessageBox.Show("Backup Directory Exists.  PLEASE IMPLEMENT ME!", "Not Implemented!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            UITimer = new Timer();
            UITimer.Interval = 100;
            UITimer.Tick += new EventHandler(UITimer_Tick);
            UITimer.Start();

			log.Debug("End Method: frmMain.frmMain()");
		}

        #endregion

#region Methods
		[STAThread]
		static void Main() 
	    {
            Application.ThreadException += (s, e) =>
            {
                log.Error("Unhandled ThreadException: " + e.Exception);
                log.Error(e.Exception.StackTrace);
            };

            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                Exception ex = (Exception)e.ExceptionObject;
                log.Error("Unhandled Exception: " + ex.Message);
                log.Error(ex.StackTrace);
            };

			Application.Run(new frmMain());
		}

		public void RefreshList()
		{
            log.Debug("Begin Method: frmMain.RefreshList()");

            ArrayList a = new ArrayList();

            foreach (string key in parser.ItemTells.Keys)
            {
                foreach (object tell in parser.ItemTells[key])
                {
                    a.Add(tell);
                }
            }

            // ItemTells is stored sorted within each matching message, with messages as the keys of the dictionary
            // If we want this behavior, then the list as it comes in by message, will be accurate.  If not, we need to resort.
            if (!chkSortIncludesItemMessage.Checked)
            {
                a.Sort();
            }

            listOfNames.Items.Clear();
			listOfTiers.Items.Clear();
			listOfDKP.Items.Clear();
            listOfAttd.Items.Clear();
            listTellType.Items.Clear();
            listItemMessage.Items.Clear();

			foreach (Raider r in a)
			{
				log.Debug("Adding Raider to list window" + r.ToString());

				listOfNames.Items.Add(r.Person);
				if (r.Tier.Equals(Raider.NOTIER)) listOfTiers.Items.Add("?");
				else listOfTiers.Items.Add(r.Tier);
				if (r.DKP == Raider.NODKP) listOfDKP.Items.Add("?");
				else listOfDKP.Items.Add(r.DKP.ToString());
                listOfAttd.Items.Add(r.AttendancePCT);
                listTellType.Items.Add(r.TellType);
                listItemMessage.Items.Add(r.ItemMessage);
			}

			log.Debug("End Method: frmMain.RefreshList()");
		}													

#endregion

#region Events
        private void UITimer_Tick(object sender, System.EventArgs e)
        {
            sbpLineCount.Text = "Lines: " + LineCount;
            sbpParseCount.Text = "Parsed: " + ParseCount;
            sbpMessage.Text = StatusMessage;
            pgbProgress.Minimum = PBMin;
            pgbProgress.Maximum = PBMax;

            try {
                if (PBVal >= PBMin && PBVal <= PBMax) pgbProgress.Value = PBVal;
            }
            catch (Exception ex) {
                log.Error("Error updating progress bar: " + ex.Message);
            }

            if (RefreshTells)
            {
                RefreshTells = false;
                this.RefreshList();
            }
        }

		private void dtpRaidDate_ValueChanged(object sender, System.EventArgs e)
		{
			log.Debug("Begin Method: frmMain.dtpRaidDate_ValueChanged(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");
			
			CurrentRaid.RaidDate = dtpRaidDate.Value;

			log.Debug("End Method: frmMain.dtpRaidDate_ValueChanged()");
		}

		private void mnuSettings_Click(object sender, System.EventArgs e)
		{
			log.Debug("Begin Method: frmMain.mnuSettings_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")"); 
			
            SettingsDialog settings = new SettingsDialog(this);
			if (settings.ShowDialog()==DialogResult.OK) 
			{
				StatusMessage = "Wrote settings to INI...";
			}
			settings.Dispose();

			log.Debug("End Method: frmMain.mnuSettings_Click()");
		}

		private void mnuDKS_Click(object sender, System.EventArgs e)
		{
			log.Debug("Begin Method: frmMain.mnuDKS_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")"); 

			TableViewer table = new TableViewer(this,TableViewer.DKPSELECT,"SELECT * FROM DKS ORDER BY Date DESC",false);
			table.Show();
            StatusMessage = "Performing SELECT query...";

			log.Debug("End Method: frmMain.mnuDKSClick()");
		}

		private void mnuAlts_Click(object sender, System.EventArgs e)
		{
			log.Debug("Begin Method: frmMain.mnuAlts_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")"); 

			TableViewer table = new TableViewer(this,TableViewer.DKPSELECT,"SELECT * FROM Alts ORDER BY AltName ASC",false);
			table.Show();
			StatusMessage = "Performing SELECT query...";

			log.Debug("End Method: frmMain.mnuDKS_Click()");
		}

		private void mnuEventLog_Click(object sender, System.EventArgs e)
		{
			log.Debug("Begin Method: frmMain.mnuEventLog_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")"); 

			TableViewer table = new TableViewer(this,TableViewer.DKPSELECT,"SELECT * FROM EventLog ORDER BY EventDate DESC",false);
			table.Show();
			StatusMessage = "Performing SELECT query...";

			log.Debug("End Method: frmMain.mnuEventLog_Click()");
		}

		private void mnuNamesTiers_Click(object sender, System.EventArgs e)
		{
			log.Debug("Begin Method: frmMain.mnuNamesTiers_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")"); 

			TableViewer table = new TableViewer(this,TableViewer.DKPSELECT,"SELECT * FROM NamesTiers ORDER BY Name ASC",false);
			table.Show();
			StatusMessage = "Performing SELECT query...";

			log.Debug("End Method: frmMain.mnuNamesTiers_Click()");
		}

		private void mnuBalanceQuery_Click(object sender, System.EventArgs e)
		{
			log.Debug("Begin Method: frmMain.mnuBalanceQuery_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			TableViewer table = new TableViewer(this,TableViewer.DKPSELECT,"SELECT Sum(DKS.PTS) AS DKPBalance, ActiveRaiders, (DKPBalance/ActiveRaiders) AS AVGBalance FROM DKS, qActiveRaiders WHERE (((DKS.Name)<>\"zzzDKP Retired\")) GROUP BY ActiveRaiders",true);
			table.Show();
			StatusMessage = "Performing SELECT query...";

			log.Debug("End Method: frmMain.mnuBalanceQuery_Click()");
		}

		private void mnuDKPTotals_Click(object sender, System.EventArgs e)
		{
			log.Debug("Begin Method: frmMain.mnuDKPTotals_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")"); 

			TableViewer table = new TableViewer(this,TableViewer.DKPSELECT,"SELECT * FROM qDKPTotals",true);
			table.Show();
			StatusMessage = "Performing SELECT query...";

			log.Debug("End Method: frmMain.mnuDKPTotals_Click()");
		}

		private void mnuRetire_Click(object sender, System.EventArgs e)
		{
			log.Debug("Begin Method: frmMain.mnuRetire_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")"); 

			InputBoxDialog input = new InputBoxDialog();
			input.FormPrompt = "Who do you want to retire? ";
			input.FormCaption = "Retire Query";
			input.ShowDialog();
			string person = input.InputResponse;
			input.Close();
			TableViewer table = new TableViewer(this,TableViewer.DKPUPDATE,"UPDATE DKS SET DKS.Name = \"zzzDKP Retired\",DKS.Comment = \"" + person + "\" WHERE (((DKS.Name)=\"" + person + "\"))",false);
			StatusMessage = "Performing UPDATE query...";

			log.Debug("End Method: frmMain.mnuRetire_Click()");
		}

		private void mnuUnretire_Click(object sender, System.EventArgs e)
		{
			log.Debug("Begin Method: frmMain.mnuUnretire_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")"); 

			InputBoxDialog input = new InputBoxDialog();
			input.FormPrompt = "Who do you want to unretire? ";
			input.FormCaption = "Unretire Query";
			input.ShowDialog();
			string person = input.InputResponse;
			input.Close();
			TableViewer table = new TableViewer(this,TableViewer.DKPUPDATE,"UPDATE DKS SET DKS.Name = \"" + person + "\",DKS.Comment = \" \" WHERE (((DKS.Name)=\"zzzDKP Retired\" AND (DKS.Comment=\"" + person + "\")))",false);
			StatusMessage = "Performing UPDATE query...";		

			log.Debug("End Method: frmMain.mnuUnretire_Click()");
		}

		private void mnuEnterSQL_Click(object sender, System.EventArgs e)
		{
			log.Debug("Begin Method: frmMain.mnuEnterSQL_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")"); 

			InputBoxDialog input = new InputBoxDialog();
			input.FormPrompt = "Enter Query: ";
			input.FormCaption = "SQL Query";
			input.ShowDialog();
			string query = input.InputResponse;
			input.Close();
			log.Debug("InputBox.InputResponse returns " + query);

			if (query.StartsWith("SELECT"))
			{
				log.Info("SELECT query entered");

				TableViewer table = new TableViewer(this,TableViewer.DKPSELECT,query,true);
				table.Show();
				StatusMessage = "Performing SELECT query...";
			}
			else if (query.StartsWith("UPDATE")||query.StartsWith("DELETE")||query.StartsWith("INSERT")) 
			{
				log.Info("UPDATE, DELETE, or INSERT query entered");

				TableViewer table = new TableViewer(this,TableViewer.DKPUPDATE,query,false);
                StatusMessage = "Performing UPDATE query...";
			}
			else
			{
				log.Info("Invalid query entered");

				MessageBox.Show("Invalid SQL Query","Error");
			}

			log.Debug("End Method: frmMain.mnuEnterSQLQuery_Click()");
		
		}

		private void mnuNameClass_Click(object sender, System.EventArgs e)
		{
			log.Debug("Begin Method: frmMain.mnuNameClass_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			TableViewer table = new TableViewer(this,TableViewer.DKPSELECT,"SELECT * FROM NamesClassesRemix ORDER BY Name",false);
			table.Show();
            StatusMessage = "Performing SELECT query...";

			log.Debug("End Method: frmMain.mnuNameClass_Click()");
		}

		private void mnuNewRaid_Click(object sender, System.EventArgs e)
		{
			log.Debug("Begin Method: frmMain.mnuNewRaid_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

            DateTime utcTime = DateTime.UtcNow;
            TimeZoneInfo serverZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
            DateTime currentDateTime = TimeZoneInfo.ConvertTimeFromUtc(utcTime, serverZone);

            CurrentRaid = new Raid(this);
            //CurrentRaid.RaidDate = System.DateTime.Today;
            CurrentRaid.RaidDate = currentDateTime.Date;
			CurrentRaid.RaidName = "";
			panel.Enabled = true;
			txtRaidName.Text = CurrentRaid.RaidName;
			dtpRaidDate.Value = CurrentRaid.RaidDate;
			chkDouble.Checked = false;
            chkSortIncludesItemMessage.Checked = false;
			txtRaidName.Enabled = true;
			dtpRaidDate.Enabled = true;
			txtZoneNames.Enabled = true;
            grpWatchFor.Enabled = true;
			grpItemTier.Enabled = true;
			listOfNames.Enabled = true;
			listOfTiers.Enabled = true;
			listOfDKP.Enabled = true;
            listOfAttd.Enabled = true;
            listTellType.Enabled = true;
            listItemMessage.Enabled = true;
            rdoB.Checked = true;

            chkWho.Checked = false;
            chkTells.Checked = false;
            chkLoot.Checked = false;

			parser.LoadNamesTiers();

			log.Info("New Raid Created.  Name: " + CurrentRaid.RaidName + " Date: " + CurrentRaid.RaidDate.ToString("dd MMM yyyy"));

			log.Debug("End Method: frmMain.mnuNewRaid_Click()");
		}

		private void btnAttendees_Click(object sender, System.EventArgs e)
		{
			log.Debug("Begin Method: frmMain.btnAttendees_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");
			
			CurrentRaid.ShowAttendees();	
	
			log.Debug("End Method: frmMain.btnAttendees_Click()");
		}

		private void mnuOpenRaid_Click(object sender, System.EventArgs e)
		{
			log.Debug("Begin Method: frmMain.mnuOpenRaid_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			ChooseRaidDialog raidselect = new ChooseRaidDialog(this);
			if (raidselect.ShowDialog() == DialogResult.OK)
			{
				CurrentRaid = new Raid(this);
				CurrentRaid.RaidDate = DateTime.Parse(raidselect.ReturnDate);
				CurrentRaid.RaidName = raidselect.ReturnRaid;
				CurrentRaid.LoadRaid();
				if (CurrentRaid.DoubleTier) chkDouble.Checked = true;
				else chkDouble.Checked = false;
				panel.Enabled = true;
				dtpRaidDate.Value = CurrentRaid.RaidDate;
				txtRaidName.Text = CurrentRaid.RaidName;
                log.Info("Raid Loaded: " + CurrentRaid.RaidName + " (" + CurrentRaid.RaidDate.ToString("dd/mm/yyyy") + ")");

                txtRaidName.Enabled = false;
				dtpRaidDate.Enabled = false;
				txtZoneNames.Enabled = true;
				grpItemTier.Enabled = true;
                grpWatchFor.Enabled = true;
				listOfNames.Enabled = true;
				listOfTiers.Enabled = true;
				listOfDKP.Enabled = true;
                listOfAttd.Enabled = true;
                listTellType.Enabled = true;
                listItemMessage.Enabled = true;

                chkWho.Checked = false;
                chkTells.Checked = false;
                chkLoot.Checked = false;

                /*
                parser.TellsOn = false;
                parser.AttendanceOn = false;
                parser.LoadNamesTiers();
                */

                rdoB.Checked = true;

				raidselect.Dispose();
			}

			log.Debug("End Method: frmMain.mnuOpenRaid_Click()");
		}

		private void btnLoot_Click(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: frmMain.btnLoot_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");
			
            CurrentRaid.ShowLoot();

            log.Debug("End Method: frmMain.btnLoot_Click()");
		}

		private void btnRecordLoot_Click(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: frmMain.btnRecordLoot_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			RecordLootDialog record = null;

			try
			{
				record = new RecordLootDialog(this,txtRaidName.Text,dtpRaidDate.Value.ToShortDateString());
			}
			catch (Exception ex)
			{
                log.Error("Could not create frmRecordLoot: " + ex.Message); 
				MessageBox.Show("Error: \n" + ex.Message);
				return;
			}
			record.ShowDialog();
			record.Dispose();
			CurrentRaid.DivideDKP();

            log.Debug("End Method: frmMain.btnRecordLoot_Click");
		}

		private void btnImport_Click(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: btnImport_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");
			if ((txtZoneNames.Text != null && txtZoneNames.Text.Length == 0)) 
			{
				MessageBox.Show("Input zone shortnames before running the attendance parser");
                log.Warn("Not loading attendance parser because no zone names are entered");
				return;
			}

			OpenFileDialog filedg = new OpenFileDialog();
			filedg.InitialDirectory = "C:\\";
			filedg.Filter = "Text File (*.txt)|*.txt|All files (*.*)|*.*";
			filedg.FilterIndex = 1;
			if(filedg.ShowDialog() == DialogResult.OK)
			{
                log.Debug("Log file dialog returns: " + filedg.FileName);
				CurrentRaid.ParseAttendance(filedg.FileName, txtZoneNames.Text);
				StatusMessage = "Finished parsing attendance...";
				filedg.Dispose();
			}

            log.Debug("End Method: btnImport_Click()");
		}

		private void txtRaidName_Leave(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: txtRaidName_Leave(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");
			
            CurrentRaid.RaidName = txtRaidName.Text;

            log.Debug("End Method: txtRaidName_Leave()");
		}

		private void chkDouble_CheckedChanged(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: chkDouble_CheckedChanged(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			CurrentRaid.DoubleTier = chkDouble.Checked;

            log.Debug("End Method: chkDouble_CheckChanged()");
		}

        private void chkSortIncludesItemMessage_CheckedChanged(object sender, EventArgs e)
        {
            RefreshTells = true;
        }

        private void btnSaveRaid_Click(object sender, System.EventArgs e)
        {
            log.Debug("Begin Method: btnSaveRaid_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			CurrentRaid.DivideDKP();
            if (CurrentRaid.Attendees == 0)
            {
                log.Warn("Not saving raid... 0 attendees");
                return;
            }
            System.Threading.Thread thd = new System.Threading.Thread(new System.Threading.ThreadStart(CurrentRaid.SyncData));
            thd.Start();
            panel.Enabled = false;
            while (thd.IsAlive)
            {
                System.Threading.Thread.Sleep(100);
                this.UITimer_Tick(this, null);
            }

            StatusMessage = "Saved raid data to database, updating tiers...";

            thd = new System.Threading.Thread(new System.Threading.ThreadStart(CurrentRaid.FigureTiers));
            thd.Start();
            while (thd.IsAlive)
            {
                System.Threading.Thread.Sleep(100);
                this.UITimer_Tick(this, null);
            }
            StatusMessage = "Updated tier standings...";
            panel.Enabled = true;

            log.Debug("End Method: btnSaveRaid_Click()");
		}

		private void btnAdd_Click(object sender, System.EventArgs e)
        {
            log.Debug("Begin Method: btnAdd_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			DataTable t = CurrentRaid.NotPresent();
			AddRemoveDialog add = new AddRemoveDialog(this,t,AddRemoveDialog.ADD,CurrentRaid.RaidName,CurrentRaid.RaidDate);
			add.ShowDialog();
			add.Dispose();

            log.Debug("End Method: btnAdd_Click()");
		}

		private void btnRemove_Click(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: btnRemove_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			DataTable t = CurrentRaid.Present();
			AddRemoveDialog remove = new AddRemoveDialog(this,t,AddRemoveDialog.REMOVE,CurrentRaid.RaidName,CurrentRaid.RaidDate);
			remove.ShowDialog();
			remove.Dispose();

            log.Debug("End Method: btnRemove_Click()");
		}

        private struct obj {
            public frmMain m;
            public DateTime t;
        }
        private obj k;

		private void mnuDailyReport_Click(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: mnuDailyReport_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			DailyReportDialog dr = new DailyReportDialog();
			DateTime t = dr.GetDate();
			if (t != DateTime.MaxValue) 
			{
                k.m = this;
                k.t = t;
                System.Threading.Thread thd = new System.Threading.Thread(new System.Threading.ThreadStart(DailyReport));
                thd.Start();
			}

            log.Debug("End Method: mnuDailyReport_Click()");
		}
        private void DailyReport()
        {
            Raid.DailyReport(k.m, k.t);
        }

		private void mnuDKPOT_Click(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: mnuDKPOT_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			DKPOverTimeDialog dkpot = new DKPOverTimeDialog(this);
			dkpot.ShowDialog();
			dkpot.Dispose();

            log.Debug("End Method: btnAdd_Click()");
		}

		private void mnuExit_Click(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: mnuExit_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			Application.Exit();

            log.Debug("End Method: mnuExit_Click()");
		}

		private void rbA_CheckedChanged(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: rbA_CheckChanged(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

            if (rdoA.Checked) this.ItemDKP = "A";
            else if (rdoB.Checked) this.ItemDKP = "B";
            else if (rdoC.Checked) this.ItemDKP = "C";
            else this.ItemDKP = "ATTD";
			
            foreach (string key in parser.ItemTells.Keys)
            {
                parser.ItemTells[key].Sort();
            }
            
			RefreshTells = true;

            log.Debug("End Method: rbA_CheckChanged()");
		}

		private void rbB_CheckedChanged(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: rbB_CheckChanged(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			if (rdoA.Checked) this.ItemDKP = "A";
			else if (rdoB.Checked) this.ItemDKP = "B";
            else if (rdoC.Checked) this.ItemDKP = "C";
            else this.ItemDKP = "ATTD";

            foreach (string key in parser.ItemTells.Keys)
            {
                parser.ItemTells[key].Sort();
            }

            RefreshTells = true;

            log.Debug("End Method: rbB_CheckChanged()");
		}

		private void rbC_CheckedChanged(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: rbC_CheckChanged(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			if (rdoA.Checked) this.ItemDKP = "A";
			else if (rdoB.Checked) this.ItemDKP = "B";
            else if (rdoC.Checked) this.ItemDKP = "C";
            else this.ItemDKP = "ATTD";

            foreach (string key in parser.ItemTells.Keys)
            {
                parser.ItemTells[key].Sort();
            }

            RefreshTells = true;

            log.Debug("End Method: rbC_CheckChanged()");
		}

        private void rbATTD_CheckedChanged(object sender, EventArgs e)
        {
            log.Debug("Begin Method: rbATTD_CheckChanged(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

            if (rdoA.Checked) this.ItemDKP = "A";
            else if (rdoB.Checked) this.ItemDKP = "B";
            else if (rdoC.Checked) this.ItemDKP = "C";
            else this.ItemDKP = "ATTD";

            foreach (string key in parser.ItemTells.Keys)
            {
                parser.ItemTells[key].Sort();
            }

            RefreshTells = true;

            log.Debug("End Method: rbATTD_CheckChanged()");
        }

		private void btnClear_Click(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: btnClear_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

            parser.ItemTells.Clear();
			RefreshTells = true;

            log.Debug("End Method: btnClear_Click()");
		}

        private void mnuAbout_Click(object sender, System.EventArgs e)
        {
            log.Debug("Begin Method: mnuAbout_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

            AboutDialog about = new AboutDialog();
            about.ShowDialog();
            about.Dispose();

            log.Debug("End Method: mnuAbout.Click()");
        }
#endregion

#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            this.mmuMain = new System.Windows.Forms.MainMenu(this.components);
            this.mnuDKP = new System.Windows.Forms.MenuItem();
            this.mnuNewRaid = new System.Windows.Forms.MenuItem();
            this.mnuOpenRaid = new System.Windows.Forms.MenuItem();
            this.mnuSettings = new System.Windows.Forms.MenuItem();
            this.mnuImport = new System.Windows.Forms.MenuItem();
            this.mnuImportLog = new System.Windows.Forms.MenuItem();
            this.mnuImportRaidDump = new System.Windows.Forms.MenuItem();
            this.mnuReports = new System.Windows.Forms.MenuItem();
            this.mnuDailyReport = new System.Windows.Forms.MenuItem();
            this.mnuDetailedTier = new System.Windows.Forms.MenuItem();
            this.mnuClassReport = new System.Windows.Forms.MenuItem();
            this.mnuDKPOT = new System.Windows.Forms.MenuItem();
            this.mnuQueries = new System.Windows.Forms.MenuItem();
            this.mnuBalanceQuery = new System.Windows.Forms.MenuItem();
            this.mnuDKPTotals = new System.Windows.Forms.MenuItem();
            this.mnuRetire = new System.Windows.Forms.MenuItem();
            this.mnuUnretire = new System.Windows.Forms.MenuItem();
            this.mnuEnterSQL = new System.Windows.Forms.MenuItem();
            this.mnuTables = new System.Windows.Forms.MenuItem();
            this.mnuDKS = new System.Windows.Forms.MenuItem();
            this.mnuAlts = new System.Windows.Forms.MenuItem();
            this.mnuEventLog = new System.Windows.Forms.MenuItem();
            this.mnuNamesTiers = new System.Windows.Forms.MenuItem();
            this.mnuNameClass = new System.Windows.Forms.MenuItem();
            this.mnuExport = new System.Windows.Forms.MenuItem();
            this.mnuExportRaidJson = new System.Windows.Forms.MenuItem();
            this.mnuExportAllRaidJson = new System.Windows.Forms.MenuItem();
            this.mnuFileSep = new System.Windows.Forms.MenuItem();
            this.mnuExit = new System.Windows.Forms.MenuItem();
            this.mnuExitNoBackup = new System.Windows.Forms.MenuItem();
            this.mnuAbout = new System.Windows.Forms.MenuItem();
            this.stbStatusBar = new System.Windows.Forms.StatusBar();
            this.sbpMessage = new System.Windows.Forms.StatusBarPanel();
            this.sbpProgressBar = new System.Windows.Forms.StatusBarPanel();
            this.sbpLineCount = new System.Windows.Forms.StatusBarPanel();
            this.sbpParseCount = new System.Windows.Forms.StatusBarPanel();
            this.panel = new System.Windows.Forms.Panel();
            this.chkSortIncludesItemMessage = new System.Windows.Forms.CheckBox();
            this.listItemMessage = new System.Windows.Forms.ListBox();
            this.lblMessage = new System.Windows.Forms.Label();
            this.listTellType = new System.Windows.Forms.ListBox();
            this.lblTellType = new System.Windows.Forms.Label();
            this.listOfAttd = new System.Windows.Forms.ListBox();
            this.listOfDKP = new System.Windows.Forms.ListBox();
            this.listOfTiers = new System.Windows.Forms.ListBox();
            this.listOfNames = new System.Windows.Forms.ListBox();
            this.txtZoneNames = new System.Windows.Forms.TextBox();
            this.grpWatchFor = new System.Windows.Forms.GroupBox();
            this.chkLoot = new System.Windows.Forms.CheckBox();
            this.chkWho = new System.Windows.Forms.CheckBox();
            this.chkTells = new System.Windows.Forms.CheckBox();
            this.lblAttd = new System.Windows.Forms.Label();
            this.btnClear = new System.Windows.Forms.Button();
            this.dtpRaidDate = new System.Windows.Forms.DateTimePicker();
            this.btnRemove = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.chkDouble = new System.Windows.Forms.CheckBox();
            this.lblZoneNames = new System.Windows.Forms.Label();
            this.btnSaveRaid = new System.Windows.Forms.Button();
            this.btnRecordLoot = new System.Windows.Forms.Button();
            this.grpItemTier = new System.Windows.Forms.GroupBox();
            this.rdoATTD = new System.Windows.Forms.RadioButton();
            this.rdoA = new System.Windows.Forms.RadioButton();
            this.rdoB = new System.Windows.Forms.RadioButton();
            this.rdoC = new System.Windows.Forms.RadioButton();
            this.dkpLabel = new System.Windows.Forms.Label();
            this.lblTier = new System.Windows.Forms.Label();
            this.lblName = new System.Windows.Forms.Label();
            this.btnLoot = new System.Windows.Forms.Button();
            this.btnAttendees = new System.Windows.Forms.Button();
            this.txtRaidName = new System.Windows.Forms.TextBox();
            this.lblRaidDate = new System.Windows.Forms.Label();
            this.lblRaidName = new System.Windows.Forms.Label();
            this.pgbProgress = new System.Windows.Forms.ProgressBar();
            ((System.ComponentModel.ISupportInitialize)(this.sbpMessage)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.sbpProgressBar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.sbpLineCount)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.sbpParseCount)).BeginInit();
            this.panel.SuspendLayout();
            this.grpWatchFor.SuspendLayout();
            this.grpItemTier.SuspendLayout();
            this.SuspendLayout();
            // 
            // mmuMain
            // 
            this.mmuMain.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mnuDKP,
            this.mnuAbout});
            // 
            // mnuDKP
            // 
            this.mnuDKP.Index = 0;
            this.mnuDKP.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mnuNewRaid,
            this.mnuOpenRaid,
            this.mnuSettings,
            this.mnuImport,
            this.mnuReports,
            this.mnuQueries,
            this.mnuTables,
            this.mnuExport,
            this.mnuFileSep,
            this.mnuExit,
            this.mnuExitNoBackup});
            this.mnuDKP.Text = "DKP";
            // 
            // mnuNewRaid
            // 
            this.mnuNewRaid.Index = 0;
            this.mnuNewRaid.Text = "New Raid";
            this.mnuNewRaid.Click += new System.EventHandler(this.mnuNewRaid_Click);
            // 
            // mnuOpenRaid
            // 
            this.mnuOpenRaid.Index = 1;
            this.mnuOpenRaid.Text = "Open Raid";
            this.mnuOpenRaid.Click += new System.EventHandler(this.mnuOpenRaid_Click);
            // 
            // mnuSettings
            // 
            this.mnuSettings.Index = 2;
            this.mnuSettings.Text = "Settings";
            this.mnuSettings.Click += new System.EventHandler(this.mnuSettings_Click);
            // 
            // mnuImport
            // 
            this.mnuImport.Index = 3;
            this.mnuImport.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mnuImportLog,
            this.mnuImportRaidDump});
            this.mnuImport.Text = "Import";
            // 
            // mnuImportLog
            // 
            this.mnuImportLog.Index = 0;
            this.mnuImportLog.Text = "Log";
            this.mnuImportLog.Click += new System.EventHandler(this.mnuImportLog_Click);
            // 
            // mnuImportRaidDump
            // 
            this.mnuImportRaidDump.Index = 1;
            this.mnuImportRaidDump.Text = "Raid Dump";
            this.mnuImportRaidDump.Click += new System.EventHandler(this.mnuImportRaidDump_Click);
            // 
            // mnuReports
            // 
            this.mnuReports.Index = 4;
            this.mnuReports.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mnuDailyReport,
            this.mnuDetailedTier,
            this.mnuClassReport,
            this.mnuDKPOT});
            this.mnuReports.Text = "Reports";
            // 
            // mnuDailyReport
            // 
            this.mnuDailyReport.Index = 0;
            this.mnuDailyReport.Text = "Daily";
            this.mnuDailyReport.Click += new System.EventHandler(this.mnuDailyReport_Click);
            // 
            // mnuDetailedTier
            // 
            this.mnuDetailedTier.Index = 1;
            this.mnuDetailedTier.Text = "Detailed Tier";
            this.mnuDetailedTier.Visible = false;
            // 
            // mnuClassReport
            // 
            this.mnuClassReport.Index = 2;
            this.mnuClassReport.Text = "Classes";
            this.mnuClassReport.Visible = false;
            // 
            // mnuDKPOT
            // 
            this.mnuDKPOT.Index = 3;
            this.mnuDKPOT.Text = "DKP Over Time";
            this.mnuDKPOT.Click += new System.EventHandler(this.mnuDKPOT_Click);
            // 
            // mnuQueries
            // 
            this.mnuQueries.Index = 5;
            this.mnuQueries.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mnuBalanceQuery,
            this.mnuDKPTotals,
            this.mnuRetire,
            this.mnuUnretire,
            this.mnuEnterSQL});
            this.mnuQueries.Text = "Queries";
            // 
            // mnuBalanceQuery
            // 
            this.mnuBalanceQuery.Index = 0;
            this.mnuBalanceQuery.Text = "DKP Balance";
            this.mnuBalanceQuery.Click += new System.EventHandler(this.mnuBalanceQuery_Click);
            // 
            // mnuDKPTotals
            // 
            this.mnuDKPTotals.Index = 1;
            this.mnuDKPTotals.Text = "DKP Totals";
            this.mnuDKPTotals.Click += new System.EventHandler(this.mnuDKPTotals_Click);
            // 
            // mnuRetire
            // 
            this.mnuRetire.Index = 2;
            this.mnuRetire.Text = "Retire";
            this.mnuRetire.Click += new System.EventHandler(this.mnuRetire_Click);
            // 
            // mnuUnretire
            // 
            this.mnuUnretire.Index = 3;
            this.mnuUnretire.Text = "Unretire";
            this.mnuUnretire.Click += new System.EventHandler(this.mnuUnretire_Click);
            // 
            // mnuEnterSQL
            // 
            this.mnuEnterSQL.Index = 4;
            this.mnuEnterSQL.Text = "Enter SQL...";
            this.mnuEnterSQL.Click += new System.EventHandler(this.mnuEnterSQL_Click);
            // 
            // mnuTables
            // 
            this.mnuTables.Index = 6;
            this.mnuTables.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mnuDKS,
            this.mnuAlts,
            this.mnuEventLog,
            this.mnuNamesTiers,
            this.mnuNameClass});
            this.mnuTables.Text = "Tables";
            // 
            // mnuDKS
            // 
            this.mnuDKS.Index = 0;
            this.mnuDKS.Text = "DKS";
            this.mnuDKS.Click += new System.EventHandler(this.mnuDKS_Click);
            // 
            // mnuAlts
            // 
            this.mnuAlts.Index = 1;
            this.mnuAlts.Text = "Alts";
            this.mnuAlts.Click += new System.EventHandler(this.mnuAlts_Click);
            // 
            // mnuEventLog
            // 
            this.mnuEventLog.Index = 2;
            this.mnuEventLog.Text = "EventLog";
            this.mnuEventLog.Click += new System.EventHandler(this.mnuEventLog_Click);
            // 
            // mnuNamesTiers
            // 
            this.mnuNamesTiers.Index = 3;
            this.mnuNamesTiers.Text = "NamesTiers";
            this.mnuNamesTiers.Click += new System.EventHandler(this.mnuNamesTiers_Click);
            // 
            // mnuNameClass
            // 
            this.mnuNameClass.Index = 4;
            this.mnuNameClass.Text = "NamesClassesRemix";
            this.mnuNameClass.Click += new System.EventHandler(this.mnuNameClass_Click);
            // 
            // mnuExport
            // 
            this.mnuExport.Index = 7;
            this.mnuExport.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
                this.mnuExportRaidJson, 
                this.mnuExportAllRaidJson
            });
            this.mnuExport.Text = "Export";
            // 
            // mnuExportRaidJson
            // 
            this.mnuExportRaidJson.Index = 0;
            this.mnuExportRaidJson.Text = "Export Raid as JSON";
            this.mnuExportRaidJson.Click += new System.EventHandler(this.mnuExportRaidJson_Click);
            // 
            // mnuExportAllRaidJson
            // 
            this.mnuExportAllRaidJson.Index = 1;
            this.mnuExportAllRaidJson.Text = "Export All Raids as JSON";
            this.mnuExportAllRaidJson.Click += new System.EventHandler(this.mnuExportAllRaidJson_Click);
            // 
            // mnuFileSep
            // 
            this.mnuFileSep.Index = 8;
            this.mnuFileSep.Text = "-";
            // 
            // mnuExit
            // 
            this.mnuExit.Index = 9;
            this.mnuExit.Text = "Exit";
            this.mnuExit.Click += new System.EventHandler(this.mnuExit_Click);
            // 
            // mnuExitNoBackup
            // 
            this.mnuExitNoBackup.Index = 10;
            this.mnuExitNoBackup.Text = "Exit w/o Backup";
            this.mnuExitNoBackup.Click += new System.EventHandler(this.mnuExitNoBackup_Click);
            // 
            // mnuAbout
            // 
            this.mnuAbout.Index = 1;
            this.mnuAbout.Text = "About";
            this.mnuAbout.Click += new System.EventHandler(this.mnuAbout_Click);
            // 
            // stbStatusBar
            // 
            this.stbStatusBar.Location = new System.Drawing.Point(0, 403);
            this.stbStatusBar.Name = "stbStatusBar";
            this.stbStatusBar.Panels.AddRange(new System.Windows.Forms.StatusBarPanel[] {
            this.sbpMessage,
            this.sbpProgressBar,
            this.sbpLineCount,
            this.sbpParseCount});
            this.stbStatusBar.ShowPanels = true;
            this.stbStatusBar.Size = new System.Drawing.Size(849, 23);
            this.stbStatusBar.SizingGrip = false;
            this.stbStatusBar.TabIndex = 0;
            // 
            // sbpMessage
            // 
            this.sbpMessage.Name = "sbpMessage";
            this.sbpMessage.Width = 279;
            // 
            // sbpProgressBar
            // 
            this.sbpProgressBar.Name = "sbpProgressBar";
            this.sbpProgressBar.Style = System.Windows.Forms.StatusBarPanelStyle.OwnerDraw;
            this.sbpProgressBar.Width = 104;
            // 
            // sbpLineCount
            // 
            this.sbpLineCount.Name = "sbpLineCount";
            this.sbpLineCount.Width = 80;
            // 
            // sbpParseCount
            // 
            this.sbpParseCount.Name = "sbpParseCount";
            this.sbpParseCount.Width = 80;
            // 
            // panel
            // 
            this.panel.Controls.Add(this.chkSortIncludesItemMessage);
            this.panel.Controls.Add(this.listItemMessage);
            this.panel.Controls.Add(this.lblMessage);
            this.panel.Controls.Add(this.listTellType);
            this.panel.Controls.Add(this.lblTellType);
            this.panel.Controls.Add(this.listOfAttd);
            this.panel.Controls.Add(this.listOfDKP);
            this.panel.Controls.Add(this.listOfTiers);
            this.panel.Controls.Add(this.listOfNames);
            this.panel.Controls.Add(this.txtZoneNames);
            this.panel.Controls.Add(this.grpWatchFor);
            this.panel.Controls.Add(this.lblAttd);
            this.panel.Controls.Add(this.btnClear);
            this.panel.Controls.Add(this.dtpRaidDate);
            this.panel.Controls.Add(this.btnRemove);
            this.panel.Controls.Add(this.btnAdd);
            this.panel.Controls.Add(this.chkDouble);
            this.panel.Controls.Add(this.lblZoneNames);
            this.panel.Controls.Add(this.btnSaveRaid);
            this.panel.Controls.Add(this.btnRecordLoot);
            this.panel.Controls.Add(this.grpItemTier);
            this.panel.Controls.Add(this.dkpLabel);
            this.panel.Controls.Add(this.lblTier);
            this.panel.Controls.Add(this.lblName);
            this.panel.Controls.Add(this.btnLoot);
            this.panel.Controls.Add(this.btnAttendees);
            this.panel.Controls.Add(this.txtRaidName);
            this.panel.Controls.Add(this.lblRaidDate);
            this.panel.Controls.Add(this.lblRaidName);
            this.panel.Enabled = false;
            this.panel.Location = new System.Drawing.Point(7, 8);
            this.panel.Name = "panel";
            this.panel.Size = new System.Drawing.Size(842, 389);
            this.panel.TabIndex = 1;
            // 
            // chkSortIncludesItemMessage
            // 
            this.chkSortIncludesItemMessage.Location = new System.Drawing.Point(132, 182);
            this.chkSortIncludesItemMessage.Name = "chkSortIncludesItemMessage";
            this.chkSortIncludesItemMessage.Size = new System.Drawing.Size(111, 23);
            this.chkSortIncludesItemMessage.TabIndex = 41;
            this.chkSortIncludesItemMessage.Text = "Sort by Message";
            this.chkSortIncludesItemMessage.CheckedChanged += new System.EventHandler(this.chkSortIncludesItemMessage_CheckedChanged);
            // 
            // listItemMessage
            // 
            this.listItemMessage.Location = new System.Drawing.Point(601, 16);
            this.listItemMessage.Name = "listItemMessage";
            this.listItemMessage.Size = new System.Drawing.Size(238, 368);
            this.listItemMessage.TabIndex = 39;
            this.listItemMessage.TabStop = false;
            // 
            // lblMessage
            // 
            this.lblMessage.Location = new System.Drawing.Point(598, 1);
            this.lblMessage.Name = "lblMessage";
            this.lblMessage.Size = new System.Drawing.Size(82, 16);
            this.lblMessage.TabIndex = 40;
            this.lblMessage.Text = "Item Message:";
            // 
            // listTellType
            // 
            this.listTellType.Location = new System.Drawing.Point(534, 16);
            this.listTellType.Name = "listTellType";
            this.listTellType.Size = new System.Drawing.Size(61, 368);
            this.listTellType.TabIndex = 37;
            this.listTellType.TabStop = false;
            // 
            // lblTellType
            // 
            this.lblTellType.Location = new System.Drawing.Point(534, 0);
            this.lblTellType.Name = "lblTellType";
            this.lblTellType.Size = new System.Drawing.Size(61, 16);
            this.lblTellType.TabIndex = 38;
            this.lblTellType.Text = "TellType:";
            // 
            // listOfAttd
            // 
            this.listOfAttd.Location = new System.Drawing.Point(491, 16);
            this.listOfAttd.Name = "listOfAttd";
            this.listOfAttd.Size = new System.Drawing.Size(37, 368);
            this.listOfAttd.TabIndex = 34;
            this.listOfAttd.TabStop = false;
            // 
            // listOfDKP
            // 
            this.listOfDKP.Location = new System.Drawing.Point(413, 16);
            this.listOfDKP.Name = "listOfDKP";
            this.listOfDKP.Size = new System.Drawing.Size(72, 368);
            this.listOfDKP.TabIndex = 14;
            this.listOfDKP.TabStop = false;
            // 
            // listOfTiers
            // 
            this.listOfTiers.Location = new System.Drawing.Point(383, 16);
            this.listOfTiers.Name = "listOfTiers";
            this.listOfTiers.Size = new System.Drawing.Size(24, 368);
            this.listOfTiers.TabIndex = 13;
            this.listOfTiers.TabStop = false;
            // 
            // listOfNames
            // 
            this.listOfNames.Location = new System.Drawing.Point(249, 16);
            this.listOfNames.Name = "listOfNames";
            this.listOfNames.Size = new System.Drawing.Size(128, 368);
            this.listOfNames.TabIndex = 12;
            this.listOfNames.TabStop = false;
            // 
            // txtZoneNames
            // 
            this.txtZoneNames.Location = new System.Drawing.Point(8, 366);
            this.txtZoneNames.Name = "txtZoneNames";
            this.txtZoneNames.Size = new System.Drawing.Size(235, 20);
            this.txtZoneNames.TabIndex = 13;
            // 
            // grpWatchFor
            // 
            this.grpWatchFor.Controls.Add(this.chkLoot);
            this.grpWatchFor.Controls.Add(this.chkWho);
            this.grpWatchFor.Controls.Add(this.chkTells);
            this.grpWatchFor.Location = new System.Drawing.Point(125, 62);
            this.grpWatchFor.Name = "grpWatchFor";
            this.grpWatchFor.Size = new System.Drawing.Size(118, 94);
            this.grpWatchFor.TabIndex = 36;
            this.grpWatchFor.TabStop = false;
            this.grpWatchFor.Text = "Watch For...";
            // 
            // chkLoot
            // 
            this.chkLoot.AutoSize = true;
            this.chkLoot.Location = new System.Drawing.Point(7, 66);
            this.chkLoot.Name = "chkLoot";
            this.chkLoot.Size = new System.Drawing.Size(47, 17);
            this.chkLoot.TabIndex = 2;
            this.chkLoot.Text = "Loot";
            this.chkLoot.UseVisualStyleBackColor = true;
            this.chkLoot.CheckedChanged += new System.EventHandler(this.chkLoot_CheckedChanged);
            // 
            // chkWho
            // 
            this.chkWho.AutoSize = true;
            this.chkWho.Location = new System.Drawing.Point(7, 43);
            this.chkWho.Name = "chkWho";
            this.chkWho.Size = new System.Drawing.Size(51, 17);
            this.chkWho.TabIndex = 1;
            this.chkWho.Text = "/who";
            this.chkWho.UseVisualStyleBackColor = true;
            this.chkWho.CheckedChanged += new System.EventHandler(this.chkWho_CheckedChanged);
            // 
            // chkTells
            // 
            this.chkTells.AutoSize = true;
            this.chkTells.Location = new System.Drawing.Point(7, 20);
            this.chkTells.Name = "chkTells";
            this.chkTells.Size = new System.Drawing.Size(48, 17);
            this.chkTells.TabIndex = 0;
            this.chkTells.Text = "Tells";
            this.chkTells.UseVisualStyleBackColor = true;
            this.chkTells.CheckedChanged += new System.EventHandler(this.chkTells_CheckedChanged);
            // 
            // lblAttd
            // 
            this.lblAttd.Location = new System.Drawing.Point(488, 1);
            this.lblAttd.Name = "lblAttd";
            this.lblAttd.Size = new System.Drawing.Size(32, 16);
            this.lblAttd.TabIndex = 35;
            this.lblAttd.Text = "Attd:";
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(8, 299);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(64, 40);
            this.btnClear.TabIndex = 33;
            this.btnClear.Text = "Clear Tells";
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // dtpRaidDate
            // 
            this.dtpRaidDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpRaidDate.Location = new System.Drawing.Point(75, 12);
            this.dtpRaidDate.Name = "dtpRaidDate";
            this.dtpRaidDate.Size = new System.Drawing.Size(128, 20);
            this.dtpRaidDate.TabIndex = 1;
            this.dtpRaidDate.Value = new System.DateTime(2006, 5, 20, 0, 0, 0, 0);
            this.dtpRaidDate.ValueChanged += new System.EventHandler(this.dtpRaidDate_ValueChanged);
            // 
            // btnRemove
            // 
            this.btnRemove.Location = new System.Drawing.Point(8, 100);
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.Size = new System.Drawing.Size(64, 32);
            this.btnRemove.TabIndex = 6;
            this.btnRemove.Text = "Remove Person";
            this.btnRemove.Click += new System.EventHandler(this.btnRemove_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(8, 62);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(64, 32);
            this.btnAdd.TabIndex = 3;
            this.btnAdd.Text = "Add Person";
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // chkDouble
            // 
            this.chkDouble.Location = new System.Drawing.Point(132, 162);
            this.chkDouble.Name = "chkDouble";
            this.chkDouble.Size = new System.Drawing.Size(88, 16);
            this.chkDouble.TabIndex = 12;
            this.chkDouble.Text = "Double Tier";
            this.chkDouble.CheckedChanged += new System.EventHandler(this.chkDouble_CheckedChanged);
            // 
            // lblZoneNames
            // 
            this.lblZoneNames.Location = new System.Drawing.Point(5, 339);
            this.lblZoneNames.Name = "lblZoneNames";
            this.lblZoneNames.Size = new System.Drawing.Size(146, 24);
            this.lblZoneNames.TabIndex = 28;
            this.lblZoneNames.Text = "Raid Zone Shortname(s):";
            this.lblZoneNames.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnSaveRaid
            // 
            this.btnSaveRaid.Location = new System.Drawing.Point(8, 252);
            this.btnSaveRaid.Name = "btnSaveRaid";
            this.btnSaveRaid.Size = new System.Drawing.Size(64, 32);
            this.btnSaveRaid.TabIndex = 10;
            this.btnSaveRaid.Text = "Save Raid";
            this.btnSaveRaid.Click += new System.EventHandler(this.btnSaveRaid_Click);
            // 
            // btnRecordLoot
            // 
            this.btnRecordLoot.Location = new System.Drawing.Point(8, 214);
            this.btnRecordLoot.Name = "btnRecordLoot";
            this.btnRecordLoot.Size = new System.Drawing.Size(64, 32);
            this.btnRecordLoot.TabIndex = 7;
            this.btnRecordLoot.Text = "Record Loot";
            this.btnRecordLoot.Click += new System.EventHandler(this.btnRecordLoot_Click);
            // 
            // grpItemTier
            // 
            this.grpItemTier.Controls.Add(this.rdoATTD);
            this.grpItemTier.Controls.Add(this.rdoA);
            this.grpItemTier.Controls.Add(this.rdoB);
            this.grpItemTier.Controls.Add(this.rdoC);
            this.grpItemTier.Location = new System.Drawing.Point(81, 299);
            this.grpItemTier.Name = "grpItemTier";
            this.grpItemTier.Size = new System.Drawing.Size(162, 48);
            this.grpItemTier.TabIndex = 22;
            this.grpItemTier.TabStop = false;
            this.grpItemTier.Text = "Item Tier";
            // 
            // rdoATTD
            // 
            this.rdoATTD.Appearance = System.Windows.Forms.Appearance.Button;
            this.rdoATTD.Location = new System.Drawing.Point(102, 16);
            this.rdoATTD.Name = "rdoATTD";
            this.rdoATTD.Size = new System.Drawing.Size(54, 24);
            this.rdoATTD.TabIndex = 17;
            this.rdoATTD.Text = "ATTD";
            this.rdoATTD.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.rdoATTD.CheckedChanged += new System.EventHandler(this.rbATTD_CheckedChanged);
            // 
            // rdoA
            // 
            this.rdoA.Appearance = System.Windows.Forms.Appearance.Button;
            this.rdoA.Location = new System.Drawing.Point(8, 16);
            this.rdoA.Name = "rdoA";
            this.rdoA.Size = new System.Drawing.Size(24, 24);
            this.rdoA.TabIndex = 14;
            this.rdoA.Text = "A";
            this.rdoA.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.rdoA.CheckedChanged += new System.EventHandler(this.rbA_CheckedChanged);
            // 
            // rdoB
            // 
            this.rdoB.Appearance = System.Windows.Forms.Appearance.Button;
            this.rdoB.Location = new System.Drawing.Point(40, 16);
            this.rdoB.Name = "rdoB";
            this.rdoB.Size = new System.Drawing.Size(24, 24);
            this.rdoB.TabIndex = 15;
            this.rdoB.Text = "B";
            this.rdoB.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.rdoB.CheckedChanged += new System.EventHandler(this.rbB_CheckedChanged);
            // 
            // rdoC
            // 
            this.rdoC.Appearance = System.Windows.Forms.Appearance.Button;
            this.rdoC.Location = new System.Drawing.Point(72, 16);
            this.rdoC.Name = "rdoC";
            this.rdoC.Size = new System.Drawing.Size(24, 24);
            this.rdoC.TabIndex = 16;
            this.rdoC.Text = "C";
            this.rdoC.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.rdoC.CheckedChanged += new System.EventHandler(this.rbC_CheckedChanged);
            // 
            // dkpLabel
            // 
            this.dkpLabel.Location = new System.Drawing.Point(410, 1);
            this.dkpLabel.Name = "dkpLabel";
            this.dkpLabel.Size = new System.Drawing.Size(56, 16);
            this.dkpLabel.TabIndex = 17;
            this.dkpLabel.Text = "DKP:";
            // 
            // lblTier
            // 
            this.lblTier.Location = new System.Drawing.Point(380, 1);
            this.lblTier.Name = "lblTier";
            this.lblTier.Size = new System.Drawing.Size(32, 16);
            this.lblTier.TabIndex = 16;
            this.lblTier.Text = "Tier:";
            // 
            // lblName
            // 
            this.lblName.Location = new System.Drawing.Point(246, 1);
            this.lblName.Name = "lblName";
            this.lblName.Size = new System.Drawing.Size(72, 16);
            this.lblName.TabIndex = 15;
            this.lblName.Text = "Name:";
            // 
            // btnLoot
            // 
            this.btnLoot.Location = new System.Drawing.Point(8, 176);
            this.btnLoot.Name = "btnLoot";
            this.btnLoot.Size = new System.Drawing.Size(64, 32);
            this.btnLoot.TabIndex = 5;
            this.btnLoot.Text = "View Loot";
            this.btnLoot.Click += new System.EventHandler(this.btnLoot_Click);
            // 
            // btnAttendees
            // 
            this.btnAttendees.Location = new System.Drawing.Point(8, 138);
            this.btnAttendees.Name = "btnAttendees";
            this.btnAttendees.Size = new System.Drawing.Size(64, 32);
            this.btnAttendees.TabIndex = 4;
            this.btnAttendees.Text = "View Attendees";
            this.btnAttendees.Click += new System.EventHandler(this.btnAttendees_Click);
            // 
            // txtRaidName
            // 
            this.txtRaidName.Location = new System.Drawing.Point(75, 36);
            this.txtRaidName.Name = "txtRaidName";
            this.txtRaidName.Size = new System.Drawing.Size(168, 20);
            this.txtRaidName.TabIndex = 2;
            this.txtRaidName.Leave += new System.EventHandler(this.txtRaidName_Leave);
            // 
            // lblRaidDate
            // 
            this.lblRaidDate.Location = new System.Drawing.Point(5, 16);
            this.lblRaidDate.Name = "lblRaidDate";
            this.lblRaidDate.Size = new System.Drawing.Size(64, 16);
            this.lblRaidDate.TabIndex = 1;
            this.lblRaidDate.Text = "Raid Date:";
            this.lblRaidDate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblRaidName
            // 
            this.lblRaidName.Location = new System.Drawing.Point(5, 37);
            this.lblRaidName.Name = "lblRaidName";
            this.lblRaidName.Size = new System.Drawing.Size(64, 16);
            this.lblRaidName.TabIndex = 0;
            this.lblRaidName.Text = "Raid Name:";
            this.lblRaidName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pgbProgress
            // 
            this.pgbProgress.Location = new System.Drawing.Point(279, 404);
            this.pgbProgress.Name = "pgbProgress";
            this.pgbProgress.Size = new System.Drawing.Size(103, 22);
            this.pgbProgress.Step = 1;
            this.pgbProgress.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.pgbProgress.TabIndex = 2;
            // 
            // frmMain
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(849, 426);
            this.Controls.Add(this.pgbProgress);
            this.Controls.Add(this.panel);
            this.Controls.Add(this.stbStatusBar);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Menu = this.mmuMain;
            this.Name = "frmMain";
            this.Text = "DKP Utilities <Eternal Sovereign>";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmMain_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.sbpMessage)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.sbpProgressBar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.sbpLineCount)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.sbpParseCount)).EndInit();
            this.panel.ResumeLayout(false);
            this.panel.PerformLayout();
            this.grpWatchFor.ResumeLayout(false);
            this.grpWatchFor.PerformLayout();
            this.grpItemTier.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

#endregion

        private void mnuImportLog_Click(object sender, EventArgs e)
        {
            log.Debug("Begin Method: mnuImportLogClick(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");
            // Do the same thing as the button.
            btnImport_Click(sender, e);
            log.Debug("End Method: mnuImportLogClick()");
        }

        private void mnuImportRaidDump_Click(object sender, EventArgs e)
        {
            log.Debug("Begin Method: mnuImportRaidDump_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

            String logdir = System.IO.Path.GetDirectoryName(this.LogFile);
            String eqdir = System.IO.Directory.GetParent(logdir).FullName;
            
            if (CurrentRaid == null)
            {
                log.Debug("Cannot parse raid attendance file if CurrentRaid is null");
                MessageBox.Show("Please have an active raid before trying to parse a raid dump file.");
            }
            else
            {
                OpenFileDialog filedg = new OpenFileDialog();
                filedg.InitialDirectory = eqdir;
                filedg.Filter = "Raid Roster (*.txt)|RaidRoster*.txt|All files (*.*)|*.*";
                filedg.FilterIndex = 1;

                if (filedg.ShowDialog() == DialogResult.OK)
                {
                    log.Debug("Log file dialog returns: " + filedg.FileName);
                    CurrentRaid.ParseRaidDump(filedg.FileName);
                    StatusMessage = "Finished parsing attendance...";
                    filedg.Dispose();
                }
            }

            log.Debug("End Method: mnuImportRaidDump_Click()");
        }

        private void chkTells_CheckedChanged(object sender, EventArgs e)
        {
            parser.TellsOn = chkTells.Checked;
        }

        private void chkWho_CheckedChanged(object sender, EventArgs e)
        {
            parser.AttendanceOn = chkWho.Checked;
        }

        private void chkLoot_CheckedChanged(object sender, EventArgs e)
        {
            parser.LootOn = chkLoot.Checked;
        }

        private void frmMain_FormClosing(object sender, EventArgs e)
        {
            if (AutomaticBackups && !(Directory.Exists(BackupDirectory))) {
                MessageBox.Show("Cannot automatically backup database - Backup Directory does not exist.  Please check your settings.",
                                "Automatic DB Backup", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else if (AutomaticBackups)
            {
                DateTime time = DateTime.UtcNow;
                Directory.CreateDirectory(BackupDirectory);
                string zipFileName = "DKP-" + time.ToString("yyyyMMdd-HHmmss") + ".zip";
                string zipFilePath = Path.Combine(BackupDirectory, zipFileName);

                string tempDirectory = Path.Combine(Directory.GetCurrentDirectory(), "temp");
                if (Directory.Exists(tempDirectory))
                {
                    Directory.Delete(tempDirectory, true);
                }
                Directory.CreateDirectory(tempDirectory);
                FileInfo dbFileInfo = new FileInfo(DBString);
                File.Copy(DBString, Path.Combine(tempDirectory, dbFileInfo.Name));
                ZipFile.CreateFromDirectory(tempDirectory, zipFilePath, CompressionLevel.Optimal, false);
                Directory.Delete(tempDirectory, true);

                MessageBox.Show("Backed up database to " + zipFilePath, "Automatic DB Backup", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void mnuExitNoBackup_Click(object sender, EventArgs e)
        {
            log.Debug("Begin Method: mnuExitNoBackup_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

            AutomaticBackups = false;
            Application.Exit();

            log.Debug("End Method: mnuExitNoBackup_Click()");
        }

        private void mnuExportRaidJson_Click(object sender, EventArgs e)
        {
            log.Debug("Begin Method: mnuExportRaidJson_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

            if (CurrentRaid != null)
            {
                CurrentRaid.SaveJsonRaidModel();
                MessageBox.Show("Successfully exported JSON Raid Model", "Success!");
            }
            else
            {
                MessageBox.Show("No raid is currently loaded", "Nope!");
            }
            

            log.Debug("End Method: mnuExportRaidJson_Click()");
        }

        private void mnuExportAllRaidJson_Click(object sender, EventArgs e)
        {
            log.Debug("Begin Method: mnuExportAllRaidJson_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

            if (CurrentRaid != null)
            {
                MessageBox.Show("There is a raid open; please close this raid or restart the program before using this feature.", "Oh No =(");
            }
            else
            {
                OleDbConnection dbConnect;
                OleDbDataAdapter dkpDA;
                DataTable raidTable = new DataTable("Raids");
                DataView raidView = new DataView(raidTable);
                string connectionStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DBString;


                try
                {
                    dbConnect = new OleDbConnection(connectionStr);
                }
                catch (Exception ex)
                {
                    log.Error("Failed to create data connection: " + ex.Message);
                    MessageBox.Show("Could not create data connection. \n(" + ex.Message + ")", "Error");
                    return;
                }

                try
                {
                    dkpDA = new OleDbDataAdapter("SELECT DISTINCT FORMAT(DKS.Date, \"yyyy/mm/dd\") AS formatDate, DKS.EventNameOrLoot AS EventName FROM DKS WHERE DKS.PTS >= 0", dbConnect);
                    dbConnect.Open();
                    dkpDA.Fill(raidTable);
                }
                catch (Exception ex)
                {
                    log.Error("Failed to get list of distinct dates and raids: " + ex.Message);
                    MessageBox.Show("Failed to get list of distinct dates and raids:\n\n" + ex.Message, "Error");
                }
                finally { dbConnect.Close(); }

                // Note: At this point, we have a list of formatDate and EventName, but there are still multiple tiers in it, such as:
                // 2023/01/12 NosSvFight_TierA0
                // 2023/01/12 NosSvFight_TierA1
                //
                // We will sort the list by formatDate and EventName, so a double tier raid should always have the raid first, and the correspoding double tier raid entry second.
                // This should allow the following pseudocode to work:
                // for row in table
                //   if raid ends in '1', skip it
                //   open raid
                //   check if next raid exists, and if it ends in '1', and set double tier to true
                //   export the raid

                raidView.Sort = "formatDate, EventName";
                int record, actualRaids = 0;

                // TODO: Make these configurable when you run the command, maybe?
                // DateTime start = DateTime.Parse(raidView[0][0].ToString());
                DateTime start = DateTime.Parse("2004/01/20");
                DateTime end = DateTime.Now;

                log.Debug($"Beginning export of {raidTable.Rows.Count} raids.");

                PBMin = 0;
                PBMax = raidTable.Rows.Count;
                PBVal = 0;

                for (record = 0; record < raidTable.Rows.Count; record++)
                {
                    if (record % 10 == 0) {
                        log.Debug($"Processed {record} out of {raidTable.Rows.Count} raid records.");
                    }

                    DataRowView row = raidView[record];

                    // If this is a "tier" record for double tier, don't process it. LoadRaid() will handle
                    if (row[1].ToString().EndsWith("1")) continue;

                    // If the raid record date doesn't fall between our start and end times, don't process it.
                    if (!DateTime.TryParse(row[0].ToString(), out DateTime raidDate) || start.Date > raidDate.Date || raidDate.Date > end.Date) continue;

                    CurrentRaid = new Raid(this);
                    CurrentRaid.RaidDate = raidDate;
                    CurrentRaid.RaidName = row[1].ToString();
                    CurrentRaid.LoadRaid();
                    if (CurrentRaid.DoubleTier) chkDouble.Checked = true;
                    else chkDouble.Checked = false;
                    panel.Enabled = true;
                    dtpRaidDate.Value = CurrentRaid.RaidDate;
                    txtRaidName.Text = CurrentRaid.RaidName;
                    log.Debug("Raid Loaded: " + CurrentRaid.RaidName + " (" + CurrentRaid.RaidDate.ToString("dd/mm/yyyy") + ")");

                    CurrentRaid.SaveJsonRaidModel();
                    actualRaids++;
                    PBVal = record;
                    sbpMessage.Text = $"Processed {record} of {raidTable.Rows.Count} raid records.";
                }

                MessageBox.Show($"Successfully exported {actualRaids} raids out of {raidTable.Rows.Count} records");
                log.Debug($"Completed export of {raidTable.Rows.Count} raids, resulting in {actualRaids} actual raid events.");
            }

            log.Debug("End Method: mnuExportAllRaidJson_Click()");
        }
    }
}
