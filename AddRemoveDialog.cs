using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

namespace ES_DKP_Utils
{

	public class AddRemoveDialog : System.Windows.Forms.Form
	{
		#region Constants
		public const int ADD = 0;
		public const int REMOVE = 1;
		#endregion

		#region Component Declarations
		private System.Windows.Forms.Label lblDate;
		private System.Windows.Forms.Label lblRaid;
		private System.Windows.Forms.ComboBox cboPeople;
		private System.Windows.Forms.Label lblAddRemove;
		private System.Windows.Forms.Button btnAddRemove;
		private System.ComponentModel.Container components = null;
		#endregion

		#region Other Declarations
		private bool add;
		private string person;
		private frmMain owner;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
		#endregion

		#region Constructor
		public AddRemoveDialog(frmMain owner, DataTable dt, int addOrRemove, string raidname, DateTime date)
		{
            log.Debug("Begin Method: frmAddRemove.frmAddRemove(frmMain,DataTable,int,string,DateTime) ("
                + owner.ToString() + "," + dt.ToString() + "," + addOrRemove.ToString() + "," + date.ToShortDateString() + ")");
			InitializeComponent();
			this.owner = owner;
			cboPeople.DataSource = dt;
			cboPeople.DisplayMember = "Name";
			lblRaid.Text = "Raid: " + raidname;
			lblDate.Text = "Date: " + date.Month + "/" + date.Day + "/" + date.Year;
			if (addOrRemove == ADD) { btnAddRemove.Text = "Add"; add = true; }
			else if (addOrRemove == REMOVE) { btnAddRemove.Text = "Remove"; add = false; }

            log.Debug("End Method: frmAddRemove.frmAddRemove()");
		}
		#endregion

		#region Events
		private void btnAddRemove_Click(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: btnAddRemove_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			if (add) 
			{
				owner.CurrentRaid.AddPerson(person);
			}
			else
			{
                owner.CurrentRaid.RemovePerson(person);
			}

            log.Debug("End Method: btnAddRemove_Click()");
		}

		private void cboPeople_Leave(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: cboPeople_Leave(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			cboPeople.SelectedIndex = cboPeople.FindString(cboPeople.Text);
            log.Debug("AutoCompleted: " + cboPeople.Text);
			person = cboPeople.Text;

            log.Debug("End Method: cboPeople_Leave()");
        }
		#endregion

		#region Windows Form Designer generated code

		private void InitializeComponent()
		{
			this.lblDate = new System.Windows.Forms.Label();
			this.lblRaid = new System.Windows.Forms.Label();
			this.cboPeople = new System.Windows.Forms.ComboBox();
			this.lblAddRemove = new System.Windows.Forms.Label();
			this.btnAddRemove = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// lblDate
			// 
			this.lblDate.Location = new System.Drawing.Point(8, 8);
			this.lblDate.Name = "lblDate";
			this.lblDate.Size = new System.Drawing.Size(120, 16);
			this.lblDate.TabIndex = 0;
			this.lblDate.Text = "Date:";
			// 
			// lblRaid
			// 
			this.lblRaid.Location = new System.Drawing.Point(8, 24);
			this.lblRaid.Name = "lblRaid";
			this.lblRaid.Size = new System.Drawing.Size(120, 16);
			this.lblRaid.TabIndex = 1;
			this.lblRaid.Text = "Raid:";
			// 
			// cboPeople
			// 
			this.cboPeople.Location = new System.Drawing.Point(16, 64);
			this.cboPeople.Name = "cboPeople";
			this.cboPeople.Size = new System.Drawing.Size(152, 21);
			this.cboPeople.TabIndex = 2;
			this.cboPeople.Leave += new System.EventHandler(this.cboPeople_Leave);
			// 
			// lblAddRemove
			// 
			this.lblAddRemove.Location = new System.Drawing.Point(0, 48);
			this.lblAddRemove.Name = "lblAddRemove";
			this.lblAddRemove.Size = new System.Drawing.Size(64, 16);
			this.lblAddRemove.TabIndex = 3;
			// 
			// btnAddRemove
			// 
			this.btnAddRemove.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.btnAddRemove.Location = new System.Drawing.Point(56, 104);
			this.btnAddRemove.Name = "btnAddRemove";
			this.btnAddRemove.Size = new System.Drawing.Size(80, 32);
			this.btnAddRemove.TabIndex = 4;
			this.btnAddRemove.Click += new System.EventHandler(this.btnAddRemove_Click);
			// 
			// frmAddRemove
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(186, 143);
			this.Controls.Add(this.btnAddRemove);
			this.Controls.Add(this.lblAddRemove);
			this.Controls.Add(this.cboPeople);
			this.Controls.Add(this.lblRaid);
			this.Controls.Add(this.lblDate);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Name = "frmAddRemove";
			this.Text = "Add / Remove Person";
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
