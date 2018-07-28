using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

namespace ES_DKP_Utils
{
	public class AppAltDialog : System.Windows.Forms.Form
	{
		#region Component Declarations
		private System.Windows.Forms.Label lblNotFound;
		private System.Windows.Forms.RadioButton rdoApp;
		private System.Windows.Forms.RadioButton rdoAlt;
		private System.Windows.Forms.Label lblApp;
		private System.Windows.Forms.Label lblAlt;
		private System.Windows.Forms.ComboBox cboPeople;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.Label lblMain;
		private System.Windows.Forms.RadioButton rdoMain;
		private System.Windows.Forms.ComboBox cboClass;
		private System.ComponentModel.Container components = null;
		#endregion

		#region Other Declarations
		private string name;
		private frmMain owner;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
		#endregion

		#region Constructor
		public AppAltDialog(frmMain owner, string name, DataTable dt)
		{         
            log.Debug("Begin Method: frmAppAlt.frmAppAlt(frmMain,string,DataTable) (" + owner.ToString() +
                "," + name.ToString() + "," + dt.ToString() + ")");

			InitializeComponent();
			this.owner = owner;
			this.name = name;
			lblNotFound.Text = name + " not found in database.  Who is s/he?";
			cboPeople.DataSource = dt;
			cboPeople.DisplayMember = "Name";
			rdoApp.Checked = true;
			cboClass.Items.AddRange( new object[] {	  "BER", "BRD", "BST", "CLR", "DRU", "ENC",
													  "MAG", "MNK", "NEC", "PAL", "RNG", "ROG",
													  "SHD", "SHM", "WAR", "WIZ" });

            log.Debug("End Method: frmAppAlt.frmAppAlt()");
		}
		#endregion

		#region Methods
		public string GetName()
		{
            log.Debug("Begin Method: frmAppAlt.GetName()");

			this.ShowDialog();
			if (this.DialogResult == DialogResult.Yes) 
			{
				owner.CurrentRaid.AddNameClass(name,cboClass.Text);
                log.Debug("End Method: frmAppAlt.GetName(), returning " + name);
				return name;
			}
            else if (this.DialogResult == DialogResult.Ignore)
            {
                log.Debug("End Method: frmAppAlt.GetName(), returning ");
                return "";
            }
            else if (this.DialogResult == DialogResult.Retry)
            {
                log.Debug("End Method: frmAppAlt.GetName(), returning " + cboPeople.Text);
                return cboPeople.Text;
            }
			return "";
		}
		#endregion

		#region Events
		private void rdoMain_CheckedChanged(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: rdoMain_CheckedChanged(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			if (rdoMain.Checked) btnOK.DialogResult = DialogResult.Yes;
			if ((cboClass.Text != null && cboClass.Text.Length == 0)) btnOK.Enabled = false;
		}

		private void rdoApp_CheckedChanged(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: rdoApp_CheckedChanged(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			if (rdoApp.Checked) btnOK.DialogResult = DialogResult.Ignore;

            log.Debug("End Method: rdoApp_CheckedChanged()");
		}

		private void rdoAlt_CheckedChanged(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: rdoAlt_CheckedChanged(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			if (rdoAlt.Checked) btnOK.DialogResult = DialogResult.Retry;

            log.Debug("End Method: rdoAlt_CheckedChanged()");
		}
		private void cboClass_Leave(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: cboClass_Leave(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			cboClass.SelectedIndex = cboClass.FindString(cboClass.Text);

            log.Debug("End Method: cboClass_Leave()");
		}
		private void frmAppAlt_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
            log.Debug("Begin Method: frmAppAlt_KeyPress(object,KeyPressEventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			if (e.KeyChar == 13) btnOK.PerformClick();

            log.Debug("End Method: frmAppAlt_KeyPress()");
		}
		
		private void cboClass_SelectedIndexChanged(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: cboClass_SelectedIndexChanged(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			if ((cboClass.Text != null && cboClass.Text.Length != 0) && cboClass.SelectedIndex != -1) btnOK.Enabled = true;

            log.Debug("End Method: cboClass_SelectedIndexChanged()");
		}
		#endregion

		#region Windows Form Designer generated code
		private void InitializeComponent()
		{
			this.lblNotFound = new System.Windows.Forms.Label();
			this.rdoApp = new System.Windows.Forms.RadioButton();
			this.rdoAlt = new System.Windows.Forms.RadioButton();
			this.lblApp = new System.Windows.Forms.Label();
			this.lblAlt = new System.Windows.Forms.Label();
			this.cboPeople = new System.Windows.Forms.ComboBox();
			this.btnOK = new System.Windows.Forms.Button();
			this.lblMain = new System.Windows.Forms.Label();
			this.rdoMain = new System.Windows.Forms.RadioButton();
			this.cboClass = new System.Windows.Forms.ComboBox();
			this.SuspendLayout();
			// 
			// lblNotFound
			// 
			this.lblNotFound.Location = new System.Drawing.Point(8, 8);
			this.lblNotFound.Name = "lblNotFound";
			this.lblNotFound.Size = new System.Drawing.Size(224, 24);
			this.lblNotFound.TabIndex = 0;
			// 
			// rdoApp
			// 
			this.rdoApp.Location = new System.Drawing.Point(24, 40);
			this.rdoApp.Name = "rdoApp";
			this.rdoApp.Size = new System.Drawing.Size(16, 16);
			this.rdoApp.TabIndex = 1;
			this.rdoApp.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.frmAppAlt_KeyPress);
			this.rdoApp.CheckedChanged += new System.EventHandler(this.rdoApp_CheckedChanged);
			// 
			// rdoAlt
			// 
			this.rdoAlt.Location = new System.Drawing.Point(24, 104);
			this.rdoAlt.Name = "rdoAlt";
			this.rdoAlt.Size = new System.Drawing.Size(16, 16);
			this.rdoAlt.TabIndex = 3;
			this.rdoAlt.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.frmAppAlt_KeyPress);
			this.rdoAlt.CheckedChanged += new System.EventHandler(this.rdoAlt_CheckedChanged);
			// 
			// lblApp
			// 
			this.lblApp.Location = new System.Drawing.Point(48, 40);
			this.lblApp.Name = "lblApp";
			this.lblApp.Size = new System.Drawing.Size(96, 16);
			this.lblApp.TabIndex = 3;
			this.lblApp.Text = "Applicant (Ignore)";
			// 
			// lblAlt
			// 
			this.lblAlt.Location = new System.Drawing.Point(48, 104);
			this.lblAlt.Name = "lblAlt";
			this.lblAlt.Size = new System.Drawing.Size(48, 16);
			this.lblAlt.TabIndex = 4;
			this.lblAlt.Text = "Alt: ";
			// 
			// cboPeople
			// 
			this.cboPeople.Location = new System.Drawing.Point(104, 104);
			this.cboPeople.Name = "cboPeople";
			this.cboPeople.Size = new System.Drawing.Size(120, 21);
			this.cboPeople.TabIndex = 5;
			// 
			// btnOK
			// 
			this.btnOK.Location = new System.Drawing.Point(64, 136);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(88, 40);
			this.btnOK.TabIndex = 6;
			this.btnOK.Text = "OK";
			// 
			// lblMain
			// 
			this.lblMain.Location = new System.Drawing.Point(48, 72);
			this.lblMain.Name = "lblMain";
			this.lblMain.Size = new System.Drawing.Size(72, 16);
			this.lblMain.TabIndex = 8;
			this.lblMain.Text = "Main (Add)";
			// 
			// rdoMain
			// 
			this.rdoMain.Location = new System.Drawing.Point(24, 72);
			this.rdoMain.Name = "rdoMain";
			this.rdoMain.Size = new System.Drawing.Size(16, 16);
			this.rdoMain.TabIndex = 2;
			this.rdoMain.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.frmAppAlt_KeyPress);
			this.rdoMain.CheckedChanged += new System.EventHandler(this.rdoMain_CheckedChanged);
			// 
			// cboClass
			// 
			this.cboClass.Location = new System.Drawing.Point(128, 72);
			this.cboClass.Name = "cboClass";
			this.cboClass.Size = new System.Drawing.Size(64, 21);
			this.cboClass.TabIndex = 4;
			this.cboClass.SelectedIndexChanged += new System.EventHandler(this.cboClass_SelectedIndexChanged);
			this.cboClass.Leave += new System.EventHandler(this.cboClass_Leave);
			// 
			// frmAppAlt
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(232, 181);
			this.Controls.Add(this.cboClass);
			this.Controls.Add(this.lblMain);
			this.Controls.Add(this.rdoMain);
			this.Controls.Add(this.btnOK);
			this.Controls.Add(this.cboPeople);
			this.Controls.Add(this.lblAlt);
			this.Controls.Add(this.lblApp);
			this.Controls.Add(this.rdoAlt);
			this.Controls.Add(this.rdoApp);
			this.Controls.Add(this.lblNotFound);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Name = "frmAppAlt";
			this.Text = "Person Not Found";
			this.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.frmAppAlt_KeyPress);
			this.ResumeLayout(false);

		}
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}
		#endregion



	}
}
