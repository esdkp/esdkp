using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using System.IO;
namespace ES_DKP_Utils
{
	public class TableViewer : System.Windows.Forms.Form
    {
        #region Constants
        public const int DKPSELECT = 0;
		public const int DKPUPDATE = 1;
        #endregion

        #region Component Declarations
        private System.Windows.Forms.DataGrid dataGrid;
		private System.ComponentModel.Container components = null;
        #endregion

        #region Other Declarations
        private OleDbConnection dbConnect;
		private OleDbDataAdapter dkpDA;
		private DataTable dbTable;
		private OleDbCommandBuilder cmdBld;
		private string query;
		private bool changed;
		private bool local;
		private frmMain owner;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public delegate void UpdateDel(DataTable dt);
        private UpdateDel update = null;
        #endregion        

        #region Constructors
        public TableViewer(frmMain owner, DataTable dt)
		{
            log.Debug("Begin Method: frmTable.frmTable(frmMain,DataTable) (" + owner.ToString() + "," + dt.ToString() + ")");

			this.owner = owner;
			local=true;
			InitializeComponent();
			dataGrid.ReadOnly = true;
			dbTable = dt;
			dataGrid.SetDataBinding(dbTable,"");
			this.dbTable.RowChanged += new DataRowChangeEventHandler(dbTable_Changed);
			this.dbTable.RowDeleted += new DataRowChangeEventHandler(dbTable_Changed);

            log.Debug("End Method: frmTable.frmTable()");
		}

		public TableViewer(frmMain owner, DataTable dt, UpdateDel u)
		{
            log.Debug("Begin Method: frmTable.frmTable(frmMain,DataTable,UpdateDel) (" + owner.ToString() + "," 
                + dt.ToString() + "," + u.ToString() + ")");

			this.owner = owner;
			local=true;
			InitializeComponent();
			dataGrid.ReadOnly = false;
			dbTable = dt;
			dataGrid.SetDataBinding(dbTable,"");
			update = u;
			this.dbTable.RowChanged += new DataRowChangeEventHandler(dbTable_Changed);
			this.dbTable.RowDeleted += new DataRowChangeEventHandler(dbTable_Changed);

            log.Debug("End Method: frmTable.frmTable()");
		}

        public TableViewer(frmMain owner, int queryType, string query, bool readOnly)
        {
            log.Debug("Begin Method: frmTable.frmTable(frmMain,int,string,bool) (" + owner.ToString() +
                "," + queryType.ToString() + "," + query.ToString() + "," + readOnly.ToString() + ")");

            this.owner = owner;
            local = false;

            if (queryType == TableViewer.DKPSELECT)
            {
                log.Debug("SELECT query initiated: " + query);

                InitializeComponent();
                changed = false;
                dataGrid.ReadOnly = readOnly;
                this.query = query;
                string connectionStr;
                connectionStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + owner.DBString;
                try { dbConnect = new OleDbConnection(connectionStr); }
                catch (Exception ex)
                {
                    log.Error("Failed to create data connection: " + ex.Message);
                    MessageBox.Show("Could not create data connection. \n(" + ex.Message + ")", "Error");
                }

                try
                {
                    this.dbTable = new DataTable("Data");
                    dkpDA = new OleDbDataAdapter(query, dbConnect);
                    dbConnect.Open();
                    dkpDA.Fill(dbTable);
                    dataGrid.SetDataBinding(dbTable, "");
                    cmdBld = new OleDbCommandBuilder(dkpDA);
                    this.dbTable.RowChanged += new DataRowChangeEventHandler(dbTable_Changed);
                    this.dbTable.RowDeleted += new DataRowChangeEventHandler(dbTable_Changed);
                }
                catch (Exception ex)
                {
                    log.Error("Failed to retrieve SELECT query: " + ex.Message);
                    MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")", "Error");
                }
                finally { dbConnect.Close(); }
            }
            else if (queryType == TableViewer.DKPUPDATE)
            {
                log.Debug("UPDATE query initiated: " + query);

                string connectionStr;
                connectionStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + owner.DBString;
                try { dbConnect = new OleDbConnection(connectionStr); }
                catch (Exception ex)
                {
                    log.Error("Failed to create data connection: " + ex.Message);
                    MessageBox.Show("Could not create data connection. \n(" + ex.Message + ")", "Error");
                }
                try
                {
                    OleDbCommand updateCommand = new OleDbCommand(query, dbConnect);
                    dbConnect.Open();
                    int changes = updateCommand.ExecuteNonQuery();
                    log.Debug("UPDATE query changed " + changes + " rows.");
                    MessageBox.Show("Updated " + changes + " rows.");

                }
                catch (Exception ex)
                {
                    log.Error("Failed to perform UPDATE query: " + ex.Message);
                    MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")", "Error");
                }
                finally { dbConnect.Close(); }
                this.Dispose();
            }

        }
        #endregion

        #region Events
        private void dbTable_Changed(object sender, System.Data.DataRowChangeEventArgs e)
		{
            log.Debug("Begin Method: dbTable_Changed(object,DataRowChangeEventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			changed = true;

            log.Debug("End Method: dbTable_Changed()");
		}

		private void frmTable_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
            log.Debug("Begin Method: frmTable_Closing(object,CancelEventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			if (changed && !dataGrid.ReadOnly && !local) 
			{
				if (MessageBox.Show("Dataset changed.  Commit to database?","Commit Changes?",MessageBoxButtons.YesNo)==DialogResult.Yes) 
				{
					try 
					{
						dkpDA.Update(dbTable);
					}
					catch (Exception ex)
					{
                        log.Error("Could not update table: " + ex.Message);

						if (MessageBox.Show("Could not save changes.\n\nError: " + ex.Message + "\n\nExit Anyway?","Error",MessageBoxButtons.YesNo) == DialogResult.No) 
						{
							e.Cancel = true;
						}
					}
				}
			} 
			else if (changed && local && !dataGrid.ReadOnly) 
			{
				if (update != null) 
				{
					update(dbTable);
				}
			}
            log.Debug("End Method: frmTable_Closing()");
		}

		private void frmTable_SizeChanged(object sender, System.EventArgs e)
		{
            log.Debug("Begin Method: frmTable_SizeChanged(object,EventArgs) (" + sender.ToString() + "," + e.ToString() + ")");

			dataGrid.Width = this.Size.Width - 6;
			dataGrid.Height = this.Size.Height - 25;

            log.Debug("End Method: frmTable_SizeChanged()");
        }
        #endregion

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(TableViewer));
            this.dataGrid = new System.Windows.Forms.DataGrid();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGrid
            // 
            this.dataGrid.AccessibleDescription = resources.GetString("dataGrid.AccessibleDescription");
            this.dataGrid.AccessibleName = resources.GetString("dataGrid.AccessibleName");
            this.dataGrid.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("dataGrid.Anchor")));
            this.dataGrid.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("dataGrid.BackgroundImage")));
            this.dataGrid.CaptionFont = ((System.Drawing.Font)(resources.GetObject("dataGrid.CaptionFont")));
            this.dataGrid.CaptionText = resources.GetString("dataGrid.CaptionText");
            this.dataGrid.DataMember = "";
            this.dataGrid.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("dataGrid.Dock")));
            this.dataGrid.Enabled = ((bool)(resources.GetObject("dataGrid.Enabled")));
            this.dataGrid.Font = ((System.Drawing.Font)(resources.GetObject("dataGrid.Font")));
            this.dataGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGrid.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("dataGrid.ImeMode")));
            this.dataGrid.Location = ((System.Drawing.Point)(resources.GetObject("dataGrid.Location")));
            this.dataGrid.Name = "dataGrid";
            this.dataGrid.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("dataGrid.RightToLeft")));
            this.dataGrid.Size = ((System.Drawing.Size)(resources.GetObject("dataGrid.Size")));
            this.dataGrid.TabIndex = ((int)(resources.GetObject("dataGrid.TabIndex")));
            this.dataGrid.Visible = ((bool)(resources.GetObject("dataGrid.Visible")));
            // 
            // frmTable
            // 
            this.AccessibleDescription = resources.GetString("$this.AccessibleDescription");
            this.AccessibleName = resources.GetString("$this.AccessibleName");
            this.AutoScaleBaseSize = ((System.Drawing.Size)(resources.GetObject("$this.AutoScaleBaseSize")));
            this.AutoScroll = ((bool)(resources.GetObject("$this.AutoScroll")));
            this.AutoScrollMargin = ((System.Drawing.Size)(resources.GetObject("$this.AutoScrollMargin")));
            this.AutoScrollMinSize = ((System.Drawing.Size)(resources.GetObject("$this.AutoScrollMinSize")));
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.ClientSize = ((System.Drawing.Size)(resources.GetObject("$this.ClientSize")));
            this.Controls.Add(this.dataGrid);
            this.Enabled = ((bool)(resources.GetObject("$this.Enabled")));
            this.Font = ((System.Drawing.Font)(resources.GetObject("$this.Font")));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("$this.ImeMode")));
            this.Location = ((System.Drawing.Point)(resources.GetObject("$this.Location")));
            this.MaximumSize = ((System.Drawing.Size)(resources.GetObject("$this.MaximumSize")));
            this.MinimumSize = ((System.Drawing.Size)(resources.GetObject("$this.MinimumSize")));
            this.Name = "frmTable";
            this.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("$this.RightToLeft")));
            this.StartPosition = ((System.Windows.Forms.FormStartPosition)(resources.GetObject("$this.StartPosition")));
            this.Text = resources.GetString("$this.Text");
            this.Closing += new System.ComponentModel.CancelEventHandler(this.frmTable_Closing);
            this.SizeChanged += new System.EventHandler(this.frmTable_SizeChanged);
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid)).EndInit();
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
