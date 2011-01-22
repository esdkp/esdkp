using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;

namespace ES_DKP_Utils
{
	public class ChooseRaidDialog : System.Windows.Forms.Form
	{
		#region Component Declarations
		private System.Windows.Forms.Label lblDate;
		private System.Windows.Forms.ComboBox cboDate;
		private System.Windows.Forms.Label lblRaid;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.ComboBox cboRaid;
		private System.ComponentModel.Container components = null;
		#endregion

		#region Other Declarations
		private OleDbConnection dbConnect;
		private OleDbDataAdapter dkpDA;
		private OleDbCommandBuilder cmdBld;
		private DataTable dateTable;
		private DataTable raidTable;
		private bool updateRaids;
		private string _ReturnRaid;
		public string ReturnRaid
		{
			get
			{
				return _ReturnRaid;
			}
			set
			{
				_ReturnRaid = value;
                
			}
		}
		private string _ReturnDate;
		public string ReturnDate
		{
			get
			{
				return _ReturnDate;
			}
			set
			{
				_ReturnDate = value;
                
			}
		}
		private frmMain owner;
        private DebugLogger debugLogger;
		#endregion

		#region Constructor
		public ChooseRaidDialog(frmMain owner)
		{
#if (DEBUG_1||DEBUG_2||DEBUG_3)
            debugLogger = new DebugLogger("frmChooseRaid.log");
#endif
            debugLogger.WriteDebug_3("Enter Method: frmChooseRaid.frmChooseRaid(frmMain) (" + owner.ToString() + ")");

			this.owner = owner;
			InitializeComponent();
			string connectionStr;
			connectionStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + owner.DBString;

			try { 
                dbConnect = new OleDbConnection(connectionStr); 
            }
			catch (Exception ex) 
			{
                debugLogger.WriteDebug_1("Failed to create data connection: " + ex.Message);
				MessageBox.Show("Could not create data connection. \n(" + ex.Message + ")","Error");
			}
			try
			{
				this.dateTable = new DataTable("Dates");
				dkpDA = new OleDbDataAdapter("SELECT DISTINCT FORMAT(DKS.Date,\"yyyy/mm/dd\") AS formatDate FROM DKS",dbConnect);
				dbConnect.Open();
				dkpDA.Fill(dateTable);
			}
			catch (Exception ex) 
			{
                debugLogger.WriteDebug_1("Failed to open list of dates: " + ex.Message);
				MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")","Error");
			}
			finally { dbConnect.Close(); }

			cboDate.DataSource = new DataView(dateTable,"","formatDate DESC",DataViewRowState.CurrentRows);
			cboDate.DisplayMember = "formatDate";

            debugLogger.WriteDebug_3("End Method: frmChooseRaid.frmChooseRaid()");
		}

		#endregion

		#region Events
		private void cboDate_SelectedIndexChanged(object sender, System.EventArgs e)
		{
            debugLogger.WriteDebug_3("Begin Method: cboDate_SelectedIndexChanged(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			updateRaids = true;

            debugLogger.WriteDebug_3("End Method: cboDate_SelectedIndexChanged()");
		}

		private void cboDate_Leave(object sender, System.EventArgs e)
		{
            debugLogger.WriteDebug_3("Begin Method: cboDate_Leave(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			if (updateRaids)
			{
				try
				{
					DateTime d = DateTime.Parse(cboDate.Text);
					raidTable = new DataTable("Raids");
					dkpDA = new OleDbDataAdapter("SELECT DISTINCT DKS.EventNameOrLoot FROM DKS WHERE (DKS.Date=#" + d.Month + "/" + d.Day + "/" + d.Year + "# AND DKS.PTS >= 0)",dbConnect);
					dbConnect.Open();
					dkpDA.Fill(raidTable);
					DataRow[] multiTiers = null;
					multiTiers = raidTable.Select("EventNameOrLoot LIKE '%0'");
					if (multiTiers.Length != 0)
					{
						foreach (DataRow r in multiTiers) 
						{
							string s = (string)r.ItemArray[0];
							s = s.Replace("0","");
							r.ItemArray = new object[1] { s };
							DataRow[] temp = raidTable.Select("EventNameOrLoot LIKE '" + s + "1'");
							foreach (DataRow dr in temp) { raidTable.Rows.Remove(dr); }
						}
					}
				}
				catch (Exception ex) 
				{
                    debugLogger.WriteDebug_1("Failed to create list of raids: " + ex.Message);
					MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")","Error");
				}
				finally { dbConnect.Close(); }
				cboRaid.DataSource = new DataView(raidTable);
				cboRaid.DisplayMember = "EventNameOrLoot";
			}

            debugLogger.WriteDebug_3("End Method: cboDate_Leave()");
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
            debugLogger.WriteDebug_3("Begin Method: btnOK_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			ReturnDate = cboDate.Text;
			ReturnRaid = cboRaid.Text;

            debugLogger.WriteDebug_3("End Method: btnOK_Click()");
		}
		#endregion

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.lblDate = new System.Windows.Forms.Label();
			this.cboDate = new System.Windows.Forms.ComboBox();
			this.cboRaid = new System.Windows.Forms.ComboBox();
			this.lblRaid = new System.Windows.Forms.Label();
			this.btnOK = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// lblDate
			// 
			this.lblDate.Location = new System.Drawing.Point(8, 16);
			this.lblDate.Name = "lblDate";
			this.lblDate.Size = new System.Drawing.Size(64, 16);
			this.lblDate.TabIndex = 0;
			this.lblDate.Text = "Raid Date:";
			// 
			// cboDate
			// 
			this.cboDate.Location = new System.Drawing.Point(80, 16);
			this.cboDate.Name = "cboDate";
			this.cboDate.Size = new System.Drawing.Size(104, 21);
			this.cboDate.TabIndex = 1;
			this.cboDate.SelectedIndexChanged += new System.EventHandler(this.cboDate_SelectedIndexChanged);
			this.cboDate.Leave += new System.EventHandler(this.cboDate_Leave);
			// 
			// cboRaid
			// 
			this.cboRaid.Location = new System.Drawing.Point(80, 48);
			this.cboRaid.Name = "cboRaid";
			this.cboRaid.Size = new System.Drawing.Size(104, 21);
			this.cboRaid.TabIndex = 2;
			// 
			// lblRaid
			// 
			this.lblRaid.Location = new System.Drawing.Point(8, 56);
			this.lblRaid.Name = "lblRaid";
			this.lblRaid.Size = new System.Drawing.Size(64, 16);
			this.lblRaid.TabIndex = 3;
			this.lblRaid.Text = "Raid:";
			// 
			// btnOK
			// 
			this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.btnOK.Location = new System.Drawing.Point(64, 96);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(72, 32);
			this.btnOK.TabIndex = 4;
			this.btnOK.Text = "OK";
			this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
			// 
			// frmChooseRaid
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(192, 141);
			this.Controls.Add(this.btnOK);
			this.Controls.Add(this.lblRaid);
			this.Controls.Add(this.cboRaid);
			this.Controls.Add(this.cboDate);
			this.Controls.Add(this.lblDate);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Name = "frmChooseRaid";
			this.Text = "Select A Raid";
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
