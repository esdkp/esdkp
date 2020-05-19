using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using Nini.Config;

namespace ES_DKP_Utils
{
	public class SettingsDialog : System.Windows.Forms.Form
	{
		#region Form Declarations
		private System.Windows.Forms.OpenFileDialog dlgFileDialog;
        private System.Windows.Forms.FolderBrowserDialog dlgFolderDialog;
		private System.Windows.Forms.TextBox txtLogFile;
		private System.Windows.Forms.Button btnChangeLogFile;
		private System.Windows.Forms.Label lblLog;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnChangeOutputDirectory;
		private System.Windows.Forms.TextBox txtOutputDirectory;
		private System.Windows.Forms.TextBox txtTax;
		private System.Windows.Forms.Label lblTax;
        private System.Windows.Forms.Label lblPct;
        private IContainer components;
        private System.Windows.Forms.Label lblReportDir;
        private Label lblDatabase;
        private Button btnChangeLocalDBFile;
        private TextBox txtLocalDBFile;
		#endregion

		#region Other Declarations
		private IConfigSource iniFile;
		private frmMain owner;
        private Label lblMinDKP;
        private TextBox txtMinDKP;
        private Label lblTierAPct;
        private TextBox txtTierAPct;
        private TextBox txtTierBPct;
        private Label lblTierBPct;
        private TextBox txtTierCPct;
        private Label lblTierCPct;
        private TextBox txtGuildNames;
        private Label lblGuildNames;
        private ToolTip toolTipGuildNames;
        private Label lblBackupDirectory;
        private TextBox txtBackupDirectory;
        private Button btnChangeBackupDirectory;
        private Label lblRaidDays;
        private TextBox txtRaidDays;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
		#endregion

		#region Constructor
		public SettingsDialog(frmMain owner)
		{
            log.Debug("Begin Method: frmSettings.frmSettings(frmMain) (" + owner.ToString() + ")");

			this.owner = owner;
			InitializeComponent();

			Directory.SetCurrentDirectory(Directory.GetParent(Application.ExecutablePath).FullName);
			iniFile = new IniConfigSource(Directory.GetCurrentDirectory() + "\\settings.ini");
			txtLogFile.Text = iniFile.Configs["Files"].GetString("logfile","");
			txtLocalDBFile.Text = iniFile.Configs["Files"].GetString("dbfile","");
			txtOutputDirectory.Text = iniFile.Configs["Files"].GetString("outdir","");
            //txtBackupDirectory.Text = iniFile.Configs["Files"].GetString("backup_directory", "\\backups");
			txtTax.Text = iniFile.Configs["Other"].GetDouble("tax",0)*100 + "";
            txtMinDKP.Text = iniFile.Configs["Other"].GetDouble("mindkp", 0) + "";
            txtTierAPct.Text = iniFile.Configs["Other"].GetDouble("tierapct", 0.60) + "";
            txtTierBPct.Text = iniFile.Configs["Other"].GetDouble("tierbpct", 0.40) + "";
            txtTierCPct.Text = iniFile.Configs["Other"].GetDouble("tiercpct", 0.30) + "";
            txtGuildNames.Text = iniFile.Configs["Other"].GetString("GuildNames", "Eternal Sovereign");
            log.Debug("End Method: frmSettings.frmSettings()");
		}
		#endregion

		#region Events
		private void btnChangeLogFile_Click(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: changeFile_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			dlgFileDialog = new OpenFileDialog();
			dlgFileDialog.InitialDirectory = "C:\\";
			dlgFileDialog.Filter = "Everquest Log File (Eqlog_*.txt)|Eqlog_*.txt|All files (*.*)|*.*";
			dlgFileDialog.FilterIndex = 1;
			if(dlgFileDialog.ShowDialog() == DialogResult.OK)
			{
                log.Debug("Select Log Dialog returns " + dlgFileDialog.FileName);

				txtLogFile.Text = dlgFileDialog.FileName;
				dlgFileDialog.Dispose();
			}

            log.Debug("End Method: changeFile_Click()");
		}

		private void btnChangeLocalDBFile_Click(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: changeDatabase_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			dlgFileDialog = new OpenFileDialog();
			dlgFileDialog.InitialDirectory = "C:\\";
			dlgFileDialog.Filter = "Microsoft Access Database (*.mdb)|*.mdb|All files (*.*)|*.*";
			dlgFileDialog.FilterIndex = 1;

			if(dlgFileDialog.ShowDialog() == DialogResult.OK)
			{
                log.Debug("Select DB Dialog returns " + dlgFileDialog.FileName);

				txtLocalDBFile.Text = dlgFileDialog.FileName;
				dlgFileDialog.Dispose();
			}

            log.Debug("End Method: changeDatabase_Click()");
		}

		private void btnChangeOutputDirectory_Click(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: changeOutput_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			dlgFolderDialog = new FolderBrowserDialog();
			dlgFolderDialog.ShowNewFolderButton = true;
			if ((txtOutputDirectory.Text != null && txtOutputDirectory.Text.Length != 0)) dlgFolderDialog.SelectedPath = txtOutputDirectory.Text;

			if (dlgFolderDialog.ShowDialog() == DialogResult.OK)
			{
                log.Debug("Select Output Directory Dialog returns " + dlgFolderDialog.SelectedPath);

				txtOutputDirectory.Text = dlgFolderDialog.SelectedPath;
				dlgFolderDialog.Dispose();
			}

            log.Debug("End Method: changeOutput_Click()");

		}

        private void btnChangeBackupDirectory_Click(object sender, System.EventArgs e)
        {
            log.Debug("Begin Method: btnChangeBackupDirectory_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

            dlgFolderDialog = new FolderBrowserDialog();
            dlgFolderDialog.ShowNewFolderButton = true;
            if ((txtBackupDirectory.Text != null && txtBackupDirectory.Text.Length != 0)) dlgFolderDialog.SelectedPath = txtBackupDirectory.Text;

            if (dlgFolderDialog.ShowDialog() == DialogResult.OK)
            {
                log.Debug("Select Backup Directory Dialog returns " + dlgFolderDialog.SelectedPath);

                txtBackupDirectory.Text = dlgFolderDialog.SelectedPath;
                dlgFolderDialog.Dispose();
            }

            log.Debug("End Method: changeOutput_Click()");

        }

        private void btnOK_Click(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: settingsOK_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			owner.DBString = txtLocalDBFile.Text;
			owner.LogFile = txtLogFile.Text;

			iniFile.Configs["Files"].Set("logfile", txtLogFile.Text);
			iniFile.Configs["Files"].Set("dbfile", txtLocalDBFile.Text);
			iniFile.Configs["Files"].Set("outdir", txtOutputDirectory.Text);
            iniFile.Configs["Files"].Set("backupdir", txtBackupDirectory.Text);
            iniFile.Configs["Other"].Set("mindkp", Double.Parse(txtMinDKP.Text));
            iniFile.Configs["Other"].Set("tierapct", Double.Parse(txtTierAPct.Text));
            iniFile.Configs["Other"].Set("tierbpct", Double.Parse(txtTierBPct.Text));
            iniFile.Configs["Other"].Set("tiercpct", Double.Parse(txtTierCPct.Text));
            iniFile.Configs["Other"].Set("GuildNames", txtGuildNames.Text);
            iniFile.Configs["Other"].Set("RaidDaysWindow", txtRaidDays.Text);

            try { 
				double d = Double.Parse(txtTax.Text);	
				iniFile.Configs["Other"].Set("tax",d/100 + "");
				iniFile.Save();
				owner.DKPTax = d/100;
                log.Info("Changed Settings: LogFile = " + 
                    owner.LogFile + ", DBFile = " + owner.DBString + ", DKPTax = " + owner.DKPTax +
                    ", OutputDirectory = " + owner.OutputDirectory);
				this.Close();
			} 
            catch (Exception ex) { 
                MessageBox.Show("Could not parse DKP tax."); 
                log.Error("Failed to parse DKP Tax: " + ex.Message); 
                this.DialogResult = DialogResult.None; 
            }

            log.Debug("End Method: settingsOK_Click()");
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: settingsCancel_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			this.Close();

            log.Debug("End Method: settingsCancel_Click()");
		}
		#endregion

        #region Windows Form Designer generated code
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingsDialog));
            this.txtLogFile = new System.Windows.Forms.TextBox();
            this.btnChangeLogFile = new System.Windows.Forms.Button();
            this.lblLog = new System.Windows.Forms.Label();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.lblReportDir = new System.Windows.Forms.Label();
            this.btnChangeOutputDirectory = new System.Windows.Forms.Button();
            this.txtOutputDirectory = new System.Windows.Forms.TextBox();
            this.txtTax = new System.Windows.Forms.TextBox();
            this.lblTax = new System.Windows.Forms.Label();
            this.lblPct = new System.Windows.Forms.Label();
            this.lblDatabase = new System.Windows.Forms.Label();
            this.btnChangeLocalDBFile = new System.Windows.Forms.Button();
            this.txtLocalDBFile = new System.Windows.Forms.TextBox();
            this.lblMinDKP = new System.Windows.Forms.Label();
            this.txtMinDKP = new System.Windows.Forms.TextBox();
            this.lblTierAPct = new System.Windows.Forms.Label();
            this.txtTierAPct = new System.Windows.Forms.TextBox();
            this.txtTierBPct = new System.Windows.Forms.TextBox();
            this.lblTierBPct = new System.Windows.Forms.Label();
            this.txtTierCPct = new System.Windows.Forms.TextBox();
            this.lblTierCPct = new System.Windows.Forms.Label();
            this.txtGuildNames = new System.Windows.Forms.TextBox();
            this.lblGuildNames = new System.Windows.Forms.Label();
            this.toolTipGuildNames = new System.Windows.Forms.ToolTip(this.components);
            this.lblRaidDays = new System.Windows.Forms.Label();
            this.lblBackupDirectory = new System.Windows.Forms.Label();
            this.txtBackupDirectory = new System.Windows.Forms.TextBox();
            this.btnChangeBackupDirectory = new System.Windows.Forms.Button();
            this.txtRaidDays = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // txtLogFile
            // 
            this.txtLogFile.Location = new System.Drawing.Point(106, 8);
            this.txtLogFile.Name = "txtLogFile";
            this.txtLogFile.Size = new System.Drawing.Size(192, 20);
            this.txtLogFile.TabIndex = 1;
            // 
            // btnChangeLogFile
            // 
            this.btnChangeLogFile.Image = ((System.Drawing.Image)(resources.GetObject("btnChangeLogFile.Image")));
            this.btnChangeLogFile.Location = new System.Drawing.Point(304, 8);
            this.btnChangeLogFile.Name = "btnChangeLogFile";
            this.btnChangeLogFile.Size = new System.Drawing.Size(24, 23);
            this.btnChangeLogFile.TabIndex = 8;
            this.btnChangeLogFile.Click += new System.EventHandler(this.btnChangeLogFile_Click);
            // 
            // lblLog
            // 
            this.lblLog.Location = new System.Drawing.Point(50, 9);
            this.lblLog.Name = "lblLog";
            this.lblLog.Size = new System.Drawing.Size(50, 16);
            this.lblLog.TabIndex = 26;
            this.lblLog.Text = "Log File:";
            this.lblLog.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Location = new System.Drawing.Point(94, 351);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(72, 32);
            this.btnOK.TabIndex = 6;
            this.btnOK.Text = "OK";
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(175, 351);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(72, 32);
            this.btnCancel.TabIndex = 7;
            this.btnCancel.Text = "Cancel";
            // 
            // lblReportDir
            // 
            this.lblReportDir.Location = new System.Drawing.Point(10, 60);
            this.lblReportDir.Name = "lblReportDir";
            this.lblReportDir.Size = new System.Drawing.Size(90, 17);
            this.lblReportDir.TabIndex = 32;
            this.lblReportDir.Text = "Report Directory:";
            this.lblReportDir.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnChangeOutputDirectory
            // 
            this.btnChangeOutputDirectory.Image = ((System.Drawing.Image)(resources.GetObject("btnChangeOutputDirectory.Image")));
            this.btnChangeOutputDirectory.Location = new System.Drawing.Point(304, 60);
            this.btnChangeOutputDirectory.Name = "btnChangeOutputDirectory";
            this.btnChangeOutputDirectory.Size = new System.Drawing.Size(24, 20);
            this.btnChangeOutputDirectory.TabIndex = 10;
            this.btnChangeOutputDirectory.Click += new System.EventHandler(this.btnChangeOutputDirectory_Click);
            // 
            // txtOutputDirectory
            // 
            this.txtOutputDirectory.Location = new System.Drawing.Point(106, 60);
            this.txtOutputDirectory.Name = "txtOutputDirectory";
            this.txtOutputDirectory.Size = new System.Drawing.Size(192, 20);
            this.txtOutputDirectory.TabIndex = 3;
            // 
            // txtTax
            // 
            this.txtTax.Location = new System.Drawing.Point(106, 125);
            this.txtTax.Name = "txtTax";
            this.txtTax.Size = new System.Drawing.Size(60, 20);
            this.txtTax.TabIndex = 4;
            // 
            // lblTax
            // 
            this.lblTax.Location = new System.Drawing.Point(44, 128);
            this.lblTax.Name = "lblTax";
            this.lblTax.Size = new System.Drawing.Size(56, 12);
            this.lblTax.TabIndex = 34;
            this.lblTax.Text = "DKP Tax:";
            this.lblTax.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblPct
            // 
            this.lblPct.Location = new System.Drawing.Point(172, 126);
            this.lblPct.Name = "lblPct";
            this.lblPct.Size = new System.Drawing.Size(15, 17);
            this.lblPct.TabIndex = 35;
            this.lblPct.Text = "%";
            this.lblPct.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblDatabase
            // 
            this.lblDatabase.Location = new System.Drawing.Point(37, 37);
            this.lblDatabase.Name = "lblDatabase";
            this.lblDatabase.Size = new System.Drawing.Size(63, 12);
            this.lblDatabase.TabIndex = 38;
            this.lblDatabase.Text = "Database:";
            this.lblDatabase.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnChangeLocalDBFile
            // 
            this.btnChangeLocalDBFile.Image = ((System.Drawing.Image)(resources.GetObject("btnChangeLocalDBFile.Image")));
            this.btnChangeLocalDBFile.Location = new System.Drawing.Point(304, 34);
            this.btnChangeLocalDBFile.Name = "btnChangeLocalDBFile";
            this.btnChangeLocalDBFile.Size = new System.Drawing.Size(24, 20);
            this.btnChangeLocalDBFile.TabIndex = 9;
            // 
            // txtLocalDBFile
            // 
            this.txtLocalDBFile.Location = new System.Drawing.Point(106, 34);
            this.txtLocalDBFile.Name = "txtLocalDBFile";
            this.txtLocalDBFile.Size = new System.Drawing.Size(192, 20);
            this.txtLocalDBFile.TabIndex = 2;
            // 
            // lblMinDKP
            // 
            this.lblMinDKP.AutoSize = true;
            this.lblMinDKP.Location = new System.Drawing.Point(27, 154);
            this.lblMinDKP.Name = "lblMinDKP";
            this.lblMinDKP.Size = new System.Drawing.Size(73, 13);
            this.lblMinDKP.TabIndex = 39;
            this.lblMinDKP.Text = "Min Tier DKP:";
            this.lblMinDKP.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtMinDKP
            // 
            this.txtMinDKP.Location = new System.Drawing.Point(106, 151);
            this.txtMinDKP.Name = "txtMinDKP";
            this.txtMinDKP.Size = new System.Drawing.Size(60, 20);
            this.txtMinDKP.TabIndex = 5;
            // 
            // lblTierAPct
            // 
            this.lblTierAPct.AutoSize = true;
            this.lblTierAPct.Location = new System.Drawing.Point(43, 181);
            this.lblTierAPct.Name = "lblTierAPct";
            this.lblTierAPct.Size = new System.Drawing.Size(57, 13);
            this.lblTierAPct.TabIndex = 40;
            this.lblTierAPct.Text = "Tier A Pct:";
            this.lblTierAPct.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtTierAPct
            // 
            this.txtTierAPct.Location = new System.Drawing.Point(106, 178);
            this.txtTierAPct.Name = "txtTierAPct";
            this.txtTierAPct.Size = new System.Drawing.Size(60, 20);
            this.txtTierAPct.TabIndex = 41;
            // 
            // txtTierBPct
            // 
            this.txtTierBPct.Location = new System.Drawing.Point(106, 201);
            this.txtTierBPct.Name = "txtTierBPct";
            this.txtTierBPct.Size = new System.Drawing.Size(60, 20);
            this.txtTierBPct.TabIndex = 43;
            // 
            // lblTierBPct
            // 
            this.lblTierBPct.AutoSize = true;
            this.lblTierBPct.Location = new System.Drawing.Point(43, 204);
            this.lblTierBPct.Name = "lblTierBPct";
            this.lblTierBPct.Size = new System.Drawing.Size(57, 13);
            this.lblTierBPct.TabIndex = 42;
            this.lblTierBPct.Text = "Tier B Pct:";
            this.lblTierBPct.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtTierCPct
            // 
            this.txtTierCPct.Location = new System.Drawing.Point(106, 227);
            this.txtTierCPct.Name = "txtTierCPct";
            this.txtTierCPct.Size = new System.Drawing.Size(60, 20);
            this.txtTierCPct.TabIndex = 45;
            // 
            // lblTierCPct
            // 
            this.lblTierCPct.AutoSize = true;
            this.lblTierCPct.Location = new System.Drawing.Point(43, 230);
            this.lblTierCPct.Name = "lblTierCPct";
            this.lblTierCPct.Size = new System.Drawing.Size(57, 13);
            this.lblTierCPct.TabIndex = 44;
            this.lblTierCPct.Text = "Tier C Pct:";
            this.lblTierCPct.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtGuildNames
            // 
            this.txtGuildNames.Location = new System.Drawing.Point(106, 254);
            this.txtGuildNames.Name = "txtGuildNames";
            this.txtGuildNames.Size = new System.Drawing.Size(192, 20);
            this.txtGuildNames.TabIndex = 46;
            // 
            // lblGuildNames
            // 
            this.lblGuildNames.AutoSize = true;
            this.lblGuildNames.Location = new System.Drawing.Point(30, 257);
            this.lblGuildNames.Name = "lblGuildNames";
            this.lblGuildNames.Size = new System.Drawing.Size(70, 13);
            this.lblGuildNames.TabIndex = 47;
            this.lblGuildNames.Text = "Guild Names:";
            this.lblGuildNames.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.toolTipGuildNames.SetToolTip(this.lblGuildNames, "Separate Guild Names with a Pipe\r\n\r\nFor example:\r\nEternal Sovereign|Crusaders Val" +
        "ourous");
            // 
            // toolTipGuildNames
            // 
            this.toolTipGuildNames.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            this.toolTipGuildNames.ToolTipTitle = "Multiple Guild Names";
            // 
            // lblRaidDays
            // 
            this.lblRaidDays.AutoSize = true;
            this.lblRaidDays.Location = new System.Drawing.Point(41, 279);
            this.lblRaidDays.Name = "lblRaidDays";
            this.lblRaidDays.Size = new System.Drawing.Size(59, 13);
            this.lblRaidDays.TabIndex = 51;
            this.lblRaidDays.Text = "Raid Days:";
            this.lblRaidDays.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.toolTipGuildNames.SetToolTip(this.lblRaidDays, "Separate Guild Names with a Pipe\r\n\r\nFor example:\r\nEternal Sovereign|Crusaders Val" +
        "ourous");
            // 
            // lblBackupDirectory
            // 
            this.lblBackupDirectory.Location = new System.Drawing.Point(-7, 84);
            this.lblBackupDirectory.Name = "lblBackupDirectory";
            this.lblBackupDirectory.Size = new System.Drawing.Size(107, 22);
            this.lblBackupDirectory.TabIndex = 48;
            this.lblBackupDirectory.Text = "Backup Directory:";
            this.lblBackupDirectory.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtBackupDirectory
            // 
            this.txtBackupDirectory.Location = new System.Drawing.Point(106, 86);
            this.txtBackupDirectory.Name = "txtBackupDirectory";
            this.txtBackupDirectory.Size = new System.Drawing.Size(192, 20);
            this.txtBackupDirectory.TabIndex = 49;
            // 
            // btnChangeBackupDirectory
            // 
            this.btnChangeBackupDirectory.Image = ((System.Drawing.Image)(resources.GetObject("btnChangeBackupDirectory.Image")));
            this.btnChangeBackupDirectory.Location = new System.Drawing.Point(304, 86);
            this.btnChangeBackupDirectory.Name = "btnChangeBackupDirectory";
            this.btnChangeBackupDirectory.Size = new System.Drawing.Size(24, 20);
            this.btnChangeBackupDirectory.TabIndex = 50;
            this.btnChangeBackupDirectory.Click += new System.EventHandler(this.btnChangeBackupDirectory_Click);
            // 
            // txtRaidDays
            // 
            this.txtRaidDays.Location = new System.Drawing.Point(106, 276);
            this.txtRaidDays.Name = "txtRaidDays";
            this.txtRaidDays.Size = new System.Drawing.Size(60, 20);
            this.txtRaidDays.TabIndex = 52;
            // 
            // SettingsDialog
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(337, 395);
            this.Controls.Add(this.txtRaidDays);
            this.Controls.Add(this.lblRaidDays);
            this.Controls.Add(this.btnChangeBackupDirectory);
            this.Controls.Add(this.txtBackupDirectory);
            this.Controls.Add(this.lblBackupDirectory);
            this.Controls.Add(this.lblGuildNames);
            this.Controls.Add(this.txtGuildNames);
            this.Controls.Add(this.txtTierCPct);
            this.Controls.Add(this.lblTierCPct);
            this.Controls.Add(this.txtTierBPct);
            this.Controls.Add(this.lblTierBPct);
            this.Controls.Add(this.txtTierAPct);
            this.Controls.Add(this.lblTierAPct);
            this.Controls.Add(this.txtMinDKP);
            this.Controls.Add(this.lblMinDKP);
            this.Controls.Add(this.lblDatabase);
            this.Controls.Add(this.btnChangeLocalDBFile);
            this.Controls.Add(this.txtLocalDBFile);
            this.Controls.Add(this.lblPct);
            this.Controls.Add(this.lblTax);
            this.Controls.Add(this.txtTax);
            this.Controls.Add(this.lblReportDir);
            this.Controls.Add(this.btnChangeOutputDirectory);
            this.Controls.Add(this.txtOutputDirectory);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.txtLogFile);
            this.Controls.Add(this.btnChangeLogFile);
            this.Controls.Add(this.lblLog);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "SettingsDialog";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "Settings";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }
        #endregion
    }
}
