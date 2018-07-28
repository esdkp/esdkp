using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace ES_DKP_Utils
{

	public class DailyReportDialog : System.Windows.Forms.Form
	{
		private System.Windows.Forms.MonthCalendar monthCalendar;
		private System.Windows.Forms.Button btnMakeReport;
		private System.ComponentModel.Container components = null;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		public DailyReportDialog()
		{
            log.Debug("Enter Method: frmDailyReport.frmDailyReport()");

			InitializeComponent();

            log.Debug("End Method: frmDailyReport.frmDailyReport()");
		}

		public DateTime GetDate()
		{
            log.Debug("Enter Method: frmDailyReport.GetDate()");

            if (this.ShowDialog() == DialogResult.OK)
            {
                log.Debug("End Method: frmDailyReport.GetDate(), returning" + monthCalendar.SelectionStart.ToShortDateString());
                return monthCalendar.SelectionStart;
            }
			this.Dispose();

            log.Debug("End Method: frmDailyReport.GetDate(), returning" + DateTime.MaxValue.ToShortDateString());
			return DateTime.MaxValue;
		}


		#region Windows Form Designer generated code
		private void InitializeComponent()
		{
			this.monthCalendar = new System.Windows.Forms.MonthCalendar();
			this.btnMakeReport = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// monthCalendar
			// 
			this.monthCalendar.Location = new System.Drawing.Point(16, 0);
			this.monthCalendar.Name = "monthCalendar";
			this.monthCalendar.ShowTodayCircle = false;
			this.monthCalendar.TabIndex = 0;
			// 
			// btnMakeReport
			// 
			this.btnMakeReport.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.btnMakeReport.Location = new System.Drawing.Point(72, 168);
			this.btnMakeReport.Name = "btnMakeReport";
			this.btnMakeReport.Size = new System.Drawing.Size(88, 32);
			this.btnMakeReport.TabIndex = 1;
			this.btnMakeReport.Text = "Make Report";
			// 
			// frmDailyReport
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(226, 215);
			this.Controls.Add(this.btnMakeReport);
			this.Controls.Add(this.monthCalendar);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Name = "frmDailyReport";
			this.Text = "Daily Report";
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
