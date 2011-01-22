using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;

namespace ES_DKP_Utils
{
    public class RecordLootDialog : System.Windows.Forms.Form
    {
        #region Component Delcarations
        private System.Windows.Forms.Label lblRaid;
        private System.Windows.Forms.Label lblRaidName;
        private System.Windows.Forms.Label lblRaidDate;
        private System.Windows.Forms.Label lblDate;
        private System.Windows.Forms.Label lblLoot;
        private System.Windows.Forms.Label lblPerson;
        private System.Windows.Forms.Label lblPrice;
        private System.Windows.Forms.ComboBox cboPeople;
        private System.Windows.Forms.ComboBox cboLoot;
        private System.Windows.Forms.ComboBox cboPrice;
        private System.Windows.Forms.Button btnRecord;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Button btnClose;
        private System.ComponentModel.Container components = null;
        #endregion

        #region Other Declarations
        private OleDbConnection dbConnect;
        private OleDbDataAdapter dkpDA;
        private OleDbCommandBuilder cmdBld;
        private DataTable dtPeople;
        private DataTable dtLoot;
        private int price;
        private int looter;
        private int loot;
        private string lootName;
        private string lootCost;
        private bool lootEntered;
        private bool nameEntered;
        private bool priceEntered;
        private frmMain owner;
        private DebugLogger debugLogger;
        #endregion

        #region Constructors
        public RecordLootDialog()
        {
#if (DEBUG_1||DEBUG_2||DEBUG_3)
            debugLogger = new DebugLogger("frmRecordLoot.log");
#endif
            debugLogger.WriteDebug_3("Begin Method: frmRecordLoot.frmRecordLoot()");

            InitializeComponent();

            debugLogger.WriteDebug_3("End Method: frmRecordLoot.frmRecordLoot()");
        }

        public RecordLootDialog(frmMain owner, string name, string date)
        {
#if (DEBUG_1||DEBUG_2||DEBUG_3)
            debugLogger = new DebugLogger("frmRecordLoot.log");
#endif

            debugLogger.WriteDebug_3("Begin Method: frmRecordLoot.frmRecordLoot(frmMain,string,string) (" + owner.ToString()
                + "," + name.ToString() + "," + date.ToString() + ")");

            InitializeComponent();

            this.owner = owner;
            lblRaidName.Text = name;
            lblRaidDate.Text = date;
            string connectionStr;
            connectionStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + owner.DBString;
            try { dbConnect = new OleDbConnection(connectionStr); }
            catch (Exception ex)
            {
                debugLogger.WriteDebug_1("Failed to open data connection: " + ex.Message);
                MessageBox.Show("Could not create data connection. \n(" + ex.Message + ")", "Error");
            }

            try
            {
                this.dtPeople = new DataTable("People");
                this.dtLoot = new DataTable("Loot");
                dkpDA = new OleDbDataAdapter("SELECT DISTINCT DKS.Name FROM DKS", dbConnect);
                dbConnect.Open();
                dkpDA.Fill(dtPeople);
                dkpDA = new OleDbDataAdapter("SELECT * FROM LootDKP ORDER BY LootName ASC", dbConnect);
                cmdBld = new OleDbCommandBuilder(dkpDA);
                dkpDA.Fill(dtLoot);
            }
            catch (Exception ex)
            {
                debugLogger.WriteDebug_1("Failed to load loot tables: " + ex.Message);
                MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")", "Error");
            }
            finally { dbConnect.Close(); }

            dtPeople.Rows.Add(new object[] { "*** Select" });
            dtLoot.Rows.Add(new object[] { "*** Select", 0 });

            DataView dvP = new DataView(dtPeople, "", "Name ASC", DataViewRowState.CurrentRows);
            DataView dvL = new DataView(dtLoot, "", "LootName ASC", DataViewRowState.CurrentRows);

            cboPrice.DataSource = dvL;
            cboPrice.DisplayMember = "DKP"; 
            cboPeople.DataSource = dvP;
            cboPeople.DisplayMember = "Name";
            cboLoot.DataSource = dvL;
            cboLoot.DisplayMember = "LootName";
            

            debugLogger.WriteDebug_3("End Method: frmRecordLoot.frmLootRecord()");
        }

        #endregion

        #region Events
        private void cboLoot_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            debugLogger.WriteDebug_3("Begin Method: cboLoot_SelectedIndexChanged(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

            string s = "Changed cboPrice.Selected Index from " + cboPrice.SelectedIndex + " to ";
            cboPrice.SelectedIndex = cboLoot.SelectedIndex;
            debugLogger.WriteDebug_2(s + cboPrice.SelectedIndex);

            debugLogger.WriteDebug_3("End Method: cboLoot_SelectedIndexChanged()");
        }
        private void cboPeople_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            debugLogger.WriteDebug_3("Begin Method: cboPeople_KeyPress(object,KeyPressEventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

            if (e.KeyChar >= 'A' && e.KeyChar <= 'z')
            {
                cboPeople.DroppedDown = true;
            }
            if (e.KeyChar == 13)
            {
                string s = "Enter key pressed in cboPeople.  cboPeople.SelectedIndex before = " + cboPeople.SelectedIndex;
                cboPeople.SelectedIndex = cboPeople.FindString(cboPeople.Text);
                s += ", after = " + cboPeople.SelectedIndex;

                if (cboPeople.SelectedIndex == -1) { cboPeople.SelectedIndex = 0; s += ". Reset to 0"; }
                debugLogger.WriteDebug_2(s);

                looter = cboPeople.SelectedIndex;
                nameEntered = true;
                cboPrice.Focus();
            }

            debugLogger.WriteDebug_3("End Method: cboPeople_KeyPress()");
        }
        private void cboPeople_Leave(object sender, System.EventArgs e)
        {
            debugLogger.WriteDebug_3("Begin Method: cboPeople_Leave(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

            if (!nameEntered)
            {
                string s = "cboPeople leaving focus, Text before = " + cboPeople.Text;
                cboPeople.SelectedIndex = cboPeople.FindString(cboPeople.Text);
                s += ", after = " + cboPeople.Text + " (" + cboPeople.SelectedIndex + ")";
                if (cboPeople.SelectedIndex == -1) { cboPeople.SelectedIndex = 0; s += " Reset to 0"; }

                debugLogger.WriteDebug_2(s);
                looter = cboPeople.SelectedIndex;
            }

            debugLogger.WriteDebug_3("End Method: cboPeople_Leave()");
        }
        private void cboLoot_Leave(object sender, System.EventArgs e)
        {
            debugLogger.WriteDebug_3("Begin Method: cboLoot_Leave(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

            if (!lootEntered)
            {
                string s = "cboLoot leaving focus, Text before = " + cboLoot.Text;
                cboLoot.SelectedIndex = cboLoot.FindString(cboLoot.Text);
                s += ", after = " + cboLoot.Text + " (" + cboLoot.SelectedIndex + ")";
                loot = cboLoot.SelectedIndex;

                debugLogger.WriteDebug_2(s);
                lootName = cboLoot.Text;

                debugLogger.WriteDebug_3("End Method: cboLoot_Leave()");
            }
        }
        private void cboLoot_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            debugLogger.WriteDebug_3("Begin Method: cboLoot_KeyPress(object,KeyPressEventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

            if (e.KeyChar >= 'A' && e.KeyChar <= 'z')
            {
                cboLoot.DroppedDown = true;
            }
            if (e.KeyChar == 13)
            {
                string s = "Enter key pressed in cboLoot.  cboLoot.Text before = " + cboLoot.Text;
                cboLoot.SelectedIndex = cboLoot.FindString(cboLoot.Text);
                s += ", after = " + cboLoot.Text + " (" + cboLoot.SelectedIndex + ")";
                loot = cboLoot.SelectedIndex;
                debugLogger.WriteDebug_2(s);

                lootName = cboLoot.Text;
                lootEntered = true;
                cboPeople.Focus();
            }

            debugLogger.WriteDebug_3("End Method: cboLoot_KeyPress()");
        }
        private void cboPrice_Leave(object sender, System.EventArgs e)
        {
            debugLogger.WriteDebug_3("Begin Method: cboPrice_Leave(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

            if (!priceEntered) lootCost = cboPrice.Text;

            debugLogger.WriteDebug_3("End Method: cboPrice_Leave()");
        }
        private void cboPrice_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            debugLogger.WriteDebug_3("Begin Method: cboPrice_KeyPress(object,KeyPressEventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

            if (e.KeyChar == 13)
            {
                debugLogger.WriteDebug_2("Enter key pressed in cboPrice, Text = " + cboPrice.Text);

                lootCost = cboPrice.Text;
                priceEntered = true;
                btnRecord.Focus();
            }

            debugLogger.WriteDebug_3("Begin Method: cboPrice_KeyPress()");
        }
        private void btnRecord_Click(object sender, System.EventArgs e)
        {
            debugLogger.WriteDebug_3("Begin Method: btnRecord_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

            debugLogger.WriteDebug_2("Upon entering btnRecord_Click, looter = " + looter + ", loot = " + loot + "," +
                   "cboPeople.Text = " + cboPeople.Text + " (" + cboPeople.SelectedIndex + "), cboLoot.Text = " + cboLoot.Text + " (" + cboLoot.SelectedIndex + ")," +
                   "cboPrice.Text = " + cboPrice.Text + " (" + cboPrice.SelectedIndex + ")");

            if (looter <= 0 || loot == 0)
            {
               MessageBox.Show("Enter name and/or loot.", "Error");
                return;
            }
            cboPeople.SelectedIndex = looter;
            cboLoot.SelectedIndex = loot;
            if (loot != -1) lootName = cboLoot.Text;
            if (lootCost == null) cboPrice_Leave(new object(), new EventArgs());

            nameEntered = lootEntered = priceEntered = false;
            DataRow r = null;
            try { price = System.Int32.Parse(lootCost); }
            catch (Exception ex) { MessageBox.Show("Invalid Price."); debugLogger.WriteDebug_1("Failed to parse price: " + ex.Message); return; }

            try { r = dtLoot.Select("LootName='" + lootName.Replace("'", "''") + "'")[0]; }
            catch (Exception ex) { debugLogger.WriteDebug_1("Failed to select loot from table: " + ex.Message);  }
            if (r == null)
            {
                debugLogger.WriteDebug_2("Item not found in Loot table: " + lootName);

                dtLoot.Rows.Add(new object[] { lootName, price });
                try
                {
                    foreach (DataRow t in dtLoot.Select("LootName='*** Select'")) dtLoot.Rows.Remove(t);
                    dkpDA.Update(dtLoot);
                }
                catch (Exception ex) { debugLogger.WriteDebug_1("Failed to update Loot table: " + ex.Message);  }
            }
            else
            {
                if ((r.ItemArray[1].GetType() == System.Type.GetType("System.DBNull")) || ((int)r["DKP"] != price))
                {
                    debugLogger.WriteDebug_2("Item price null or different in Loot table, updating");
                    try
                    {
                        r["DKP"] = price;
                        foreach (DataRow t in dtLoot.Select("LootName='*** Select'")) dtLoot.Rows.Remove(t);
                        dkpDA.Update(dtLoot);
                    }
                    catch (Exception ex) { debugLogger.WriteDebug_1("Failed to update Loot table: " + ex.Message); }
                }
            }
            try
            {
                int cost = -1 * price;
                object[] o = new object[] { cboPeople.Text, "#" + lblRaidDate.Text + "#", lootName, cost.ToString(), null, lblRaidName.Text };
                owner.CurrentRaid.AddRowToLocalTable(o);

                lblStatus.Text = "Awarded " + cboPeople.Text + " " + lootName + " for a cost of " + price + " DKP.";
            }
            catch (Exception ex) { debugLogger.WriteDebug_1("Failed to add loot row: " + ex.Message); }
        }
        private void btnClose_Click(object sender, System.EventArgs e)
        {
            debugLogger.WriteDebug_3("Begin Method: btnClose_Click(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

            this.Close();

            debugLogger.WriteDebug_3("Begin Method: btnClose_Click()");
        }

        #endregion

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lblRaid = new System.Windows.Forms.Label();
            this.lblRaidName = new System.Windows.Forms.Label();
            this.lblRaidDate = new System.Windows.Forms.Label();
            this.lblDate = new System.Windows.Forms.Label();
            this.lblLoot = new System.Windows.Forms.Label();
            this.lblPerson = new System.Windows.Forms.Label();
            this.lblPrice = new System.Windows.Forms.Label();
            this.cboPeople = new System.Windows.Forms.ComboBox();
            this.cboLoot = new System.Windows.Forms.ComboBox();
            this.cboPrice = new System.Windows.Forms.ComboBox();
            this.btnRecord = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.btnClose = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblRaid
            // 
            this.lblRaid.Location = new System.Drawing.Point(8, 16);
            this.lblRaid.Name = "lblRaid";
            this.lblRaid.Size = new System.Drawing.Size(32, 16);
            this.lblRaid.TabIndex = 0;
            this.lblRaid.Text = "Raid:";
            // 
            // lblRaidName
            // 
            this.lblRaidName.Font = new System.Drawing.Font("Comic Sans MS", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblRaidName.Location = new System.Drawing.Point(56, 16);
            this.lblRaidName.Name = "lblRaidName";
            this.lblRaidName.Size = new System.Drawing.Size(112, 24);
            this.lblRaidName.TabIndex = 1;
            // 
            // lblRaidDate
            // 
            this.lblRaidDate.Font = new System.Drawing.Font("Comic Sans MS", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblRaidDate.Location = new System.Drawing.Point(56, 40);
            this.lblRaidDate.Name = "lblRaidDate";
            this.lblRaidDate.Size = new System.Drawing.Size(112, 24);
            this.lblRaidDate.TabIndex = 3;
            // 
            // lblDate
            // 
            this.lblDate.Location = new System.Drawing.Point(8, 40);
            this.lblDate.Name = "lblDate";
            this.lblDate.Size = new System.Drawing.Size(32, 16);
            this.lblDate.TabIndex = 2;
            this.lblDate.Text = "Date:";
            // 
            // lblLoot
            // 
            this.lblLoot.Location = new System.Drawing.Point(8, 64);
            this.lblLoot.Name = "lblLoot";
            this.lblLoot.Size = new System.Drawing.Size(32, 24);
            this.lblLoot.TabIndex = 4;
            this.lblLoot.Text = "Loot: ";
            // 
            // lblPerson
            // 
            this.lblPerson.Location = new System.Drawing.Point(8, 96);
            this.lblPerson.Name = "lblPerson";
            this.lblPerson.Size = new System.Drawing.Size(48, 24);
            this.lblPerson.TabIndex = 5;
            this.lblPerson.Text = "Person:";
            // 
            // lblPrice
            // 
            this.lblPrice.Location = new System.Drawing.Point(8, 128);
            this.lblPrice.Name = "lblPrice";
            this.lblPrice.Size = new System.Drawing.Size(40, 24);
            this.lblPrice.TabIndex = 6;
            this.lblPrice.Text = "Price:";
            // 
            // cboPeople
            // 
            this.cboPeople.CausesValidation = false;
            this.cboPeople.Location = new System.Drawing.Point(56, 96);
            this.cboPeople.Name = "cboPeople";
            this.cboPeople.Size = new System.Drawing.Size(176, 21);
            this.cboPeople.TabIndex = 2;
            this.cboPeople.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cboPeople_KeyPress);
            this.cboPeople.Leave += new System.EventHandler(this.cboPeople_Leave);
            // 
            // cboLoot
            // 
            this.cboLoot.DisplayMember = "LootName";
            this.cboLoot.Location = new System.Drawing.Point(56, 64);
            this.cboLoot.Name = "cboLoot";
            this.cboLoot.Size = new System.Drawing.Size(176, 21);
            this.cboLoot.TabIndex = 1;
            this.cboLoot.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cboLoot_KeyPress);
            this.cboLoot.SelectedIndexChanged += new System.EventHandler(this.cboLoot_SelectedIndexChanged);
            this.cboLoot.Leave += new System.EventHandler(this.cboLoot_Leave);
            // 
            // cboPrice
            // 
            this.cboPrice.Location = new System.Drawing.Point(56, 128);
            this.cboPrice.Name = "cboPrice";
            this.cboPrice.Size = new System.Drawing.Size(176, 21);
            this.cboPrice.TabIndex = 3;
            this.cboPrice.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cboPrice_KeyPress);
            this.cboPrice.Leave += new System.EventHandler(this.cboPrice_Leave);
            // 
            // btnRecord
            // 
            this.btnRecord.Location = new System.Drawing.Point(160, 152);
            this.btnRecord.Name = "btnRecord";
            this.btnRecord.Size = new System.Drawing.Size(72, 32);
            this.btnRecord.TabIndex = 4;
            this.btnRecord.Text = "Record";
            this.btnRecord.Click += new System.EventHandler(this.btnRecord_Click);
            // 
            // lblStatus
            // 
            this.lblStatus.Location = new System.Drawing.Point(8, 160);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(144, 64);
            this.lblStatus.TabIndex = 7;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(160, 192);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(72, 32);
            this.btnClose.TabIndex = 5;
            this.btnClose.Text = "Close";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // frmRecordLoot
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(232, 231);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnRecord);
            this.Controls.Add(this.cboPrice);
            this.Controls.Add(this.cboLoot);
            this.Controls.Add(this.cboPeople);
            this.Controls.Add(this.lblPrice);
            this.Controls.Add(this.lblPerson);
            this.Controls.Add(this.lblLoot);
            this.Controls.Add(this.lblRaidDate);
            this.Controls.Add(this.lblDate);
            this.Controls.Add(this.lblRaidName);
            this.Controls.Add(this.lblRaid);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "frmRecordLoot";
            this.Text = "Record loot";
            this.ResumeLayout(false);

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
