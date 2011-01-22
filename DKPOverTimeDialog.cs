using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Drawing.Text;


namespace ES_DKP_Utils
{

	public class DKPOverTimeDialog : System.Windows.Forms.Form
	{
		#region Settings
		private Color bgColor = Color.Gray;
		private Color axisColor = Color.Black;
		private Color crossNetColor = Color.Purple;
		private Color crossIncColor = Color.Blue;
		private Color crossOutColor = Color.Red;
		private Color lineNetColor = Color.HotPink;
		private Color lineIncColor = Color.Green;
		private Color lineOutColor = Color.Orange;
		private Color textColor = Color.Black;
		private int pointWidth = 1;
		private int crossLength = 3;
		private int lineWidth = 2;
		private int axisWidth = 2;
		private Font labelFont = new Font("Courier New",12);
		private Font labelFontSmall = new Font("Arial Narrow",10);
		private int xlabelspace = 75; // 75px between labels
		private int ylabelspace = 50;		
		private int  breakpt = 0;
		private bool retired = false;
		#endregion

		#region Form Components
		private System.Windows.Forms.DateTimePicker dateTimePicker1;
		private System.Windows.Forms.DateTimePicker dateTimePicker2;
		private System.Windows.Forms.Label lblStart;
		private System.Windows.Forms.Label lblEnd;
		private System.Windows.Forms.Label lblBreak;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox grpOne;
		private System.Windows.Forms.Label lblBRK1;
		private System.Windows.Forms.Label lblBRK2;
		private System.Windows.Forms.Button btnGenerate;
		private System.Windows.Forms.RadioButton rdoSingle;
		private System.Windows.Forms.RadioButton rdoMulti;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.RadioButton rdoTotal;
		private System.Windows.Forms.Label lblTotal;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.NumericUpDown nmBreakpt;
		private System.Windows.Forms.TextBox txtXSize;
		private System.Windows.Forms.TextBox txtYSize;
		private System.Windows.Forms.Label lblSX;
		private System.Windows.Forms.Label lblSY;
		private System.Windows.Forms.CheckBox chkInc;
		private System.Windows.Forms.CheckBox chkOut;
		private System.Windows.Forms.CheckBox chkNet;
		private System.Windows.Forms.Label lblNET;
		private System.Windows.Forms.Label lblOUT;
		private System.Windows.Forms.Label lblINC;
		private System.Windows.Forms.RadioButton rdoChange;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.RadioButton rdoWeek;
		private System.Windows.Forms.RadioButton rdoDay;
		private System.Windows.Forms.RadioButton rdoMonth;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.GroupBox grpRet;
		private System.Windows.Forms.RadioButton rdoRetInc;
		private System.Windows.Forms.RadioButton rdoRetExc;
		private System.Windows.Forms.Label lblRetInc;
		private System.Windows.Forms.Label lblRetExc;
		private System.Windows.Forms.CheckBox chkPts;
		private System.Windows.Forms.Label lblPoints;
		private System.Windows.Forms.Label lblLines;
		private System.Windows.Forms.CheckBox chkLines;
		private System.ComponentModel.Container components = null;
		#endregion

		#region Other Components
		private frmMain parent;
		#endregion

		#region Constructor
		public DKPOverTimeDialog(frmMain parent)
		{
			InitializeComponent();
			this.parent = parent;
		}
		#endregion

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.btnGenerate = new System.Windows.Forms.Button();
			this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
			this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
			this.lblStart = new System.Windows.Forms.Label();
			this.lblEnd = new System.Windows.Forms.Label();
			this.rdoSingle = new System.Windows.Forms.RadioButton();
			this.rdoMulti = new System.Windows.Forms.RadioButton();
			this.lblBreak = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.grpOne = new System.Windows.Forms.GroupBox();
			this.lblSX = new System.Windows.Forms.Label();
			this.txtYSize = new System.Windows.Forms.TextBox();
			this.txtXSize = new System.Windows.Forms.TextBox();
			this.lblBRK2 = new System.Windows.Forms.Label();
			this.lblBRK1 = new System.Windows.Forms.Label();
			this.nmBreakpt = new System.Windows.Forms.NumericUpDown();
			this.lblSY = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblLines = new System.Windows.Forms.Label();
			this.chkLines = new System.Windows.Forms.CheckBox();
			this.lblPoints = new System.Windows.Forms.Label();
			this.chkPts = new System.Windows.Forms.CheckBox();
			this.lblINC = new System.Windows.Forms.Label();
			this.lblOUT = new System.Windows.Forms.Label();
			this.lblNET = new System.Windows.Forms.Label();
			this.chkNet = new System.Windows.Forms.CheckBox();
			this.chkOut = new System.Windows.Forms.CheckBox();
			this.chkInc = new System.Windows.Forms.CheckBox();
			this.label2 = new System.Windows.Forms.Label();
			this.rdoChange = new System.Windows.Forms.RadioButton();
			this.lblTotal = new System.Windows.Forms.Label();
			this.rdoTotal = new System.Windows.Forms.RadioButton();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label5 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.rdoMonth = new System.Windows.Forms.RadioButton();
			this.rdoDay = new System.Windows.Forms.RadioButton();
			this.rdoWeek = new System.Windows.Forms.RadioButton();
			this.grpRet = new System.Windows.Forms.GroupBox();
			this.lblRetExc = new System.Windows.Forms.Label();
			this.lblRetInc = new System.Windows.Forms.Label();
			this.rdoRetExc = new System.Windows.Forms.RadioButton();
			this.rdoRetInc = new System.Windows.Forms.RadioButton();
			this.grpOne.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.nmBreakpt)).BeginInit();
			this.groupBox1.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.grpRet.SuspendLayout();
			this.SuspendLayout();
			// 
			// btnGenerate
			// 
			this.btnGenerate.Location = new System.Drawing.Point(184, 264);
			this.btnGenerate.Name = "btnGenerate";
			this.btnGenerate.Size = new System.Drawing.Size(88, 40);
			this.btnGenerate.TabIndex = 0;
			this.btnGenerate.Text = "Generate";
			this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
			// 
			// dateTimePicker1
			// 
			this.dateTimePicker1.Location = new System.Drawing.Point(96, 8);
			this.dateTimePicker1.Name = "dateTimePicker1";
			this.dateTimePicker1.Size = new System.Drawing.Size(192, 20);
			this.dateTimePicker1.TabIndex = 1;
			// 
			// dateTimePicker2
			// 
			this.dateTimePicker2.Location = new System.Drawing.Point(96, 40);
			this.dateTimePicker2.Name = "dateTimePicker2";
			this.dateTimePicker2.Size = new System.Drawing.Size(192, 20);
			this.dateTimePicker2.TabIndex = 2;
			// 
			// lblStart
			// 
			this.lblStart.Location = new System.Drawing.Point(40, 8);
			this.lblStart.Name = "lblStart";
			this.lblStart.Size = new System.Drawing.Size(32, 24);
			this.lblStart.TabIndex = 3;
			this.lblStart.Text = "Start:";
			// 
			// lblEnd
			// 
			this.lblEnd.Location = new System.Drawing.Point(40, 40);
			this.lblEnd.Name = "lblEnd";
			this.lblEnd.Size = new System.Drawing.Size(32, 24);
			this.lblEnd.TabIndex = 4;
			this.lblEnd.Text = "End:";
			// 
			// rdoSingle
			// 
			this.rdoSingle.Checked = true;
			this.rdoSingle.Location = new System.Drawing.Point(8, 24);
			this.rdoSingle.Name = "rdoSingle";
			this.rdoSingle.Size = new System.Drawing.Size(16, 8);
			this.rdoSingle.TabIndex = 5;
			this.rdoSingle.TabStop = true;
			this.rdoSingle.CheckedChanged += new System.EventHandler(this.rdoSingle_CheckedChanged);
			// 
			// rdoMulti
			// 
			this.rdoMulti.Location = new System.Drawing.Point(8, 40);
			this.rdoMulti.Name = "rdoMulti";
			this.rdoMulti.Size = new System.Drawing.Size(16, 22);
			this.rdoMulti.TabIndex = 7;
			this.rdoMulti.CheckedChanged += new System.EventHandler(this.rdoMulti_CheckedChanged);
			// 
			// lblBreak
			// 
			this.lblBreak.Location = new System.Drawing.Point(24, 24);
			this.lblBreak.Name = "lblBreak";
			this.lblBreak.Size = new System.Drawing.Size(64, 16);
			this.lblBreak.TabIndex = 8;
			this.lblBreak.Text = "One Image";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(24, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(88, 16);
			this.label1.TabIndex = 9;
			this.label1.Text = "Multiple Images";
			// 
			// grpOne
			// 
			this.grpOne.Controls.Add(this.lblSX);
			this.grpOne.Controls.Add(this.txtYSize);
			this.grpOne.Controls.Add(this.txtXSize);
			this.grpOne.Controls.Add(this.lblBRK2);
			this.grpOne.Controls.Add(this.lblBRK1);
			this.grpOne.Controls.Add(this.nmBreakpt);
			this.grpOne.Controls.Add(this.label1);
			this.grpOne.Controls.Add(this.rdoSingle);
			this.grpOne.Controls.Add(this.rdoMulti);
			this.grpOne.Controls.Add(this.lblBreak);
			this.grpOne.Controls.Add(this.lblSY);
			this.grpOne.Location = new System.Drawing.Point(16, 72);
			this.grpOne.Name = "grpOne";
			this.grpOne.Size = new System.Drawing.Size(160, 136);
			this.grpOne.TabIndex = 13;
			this.grpOne.TabStop = false;
			this.grpOne.Text = "Output";
			// 
			// lblSX
			// 
			this.lblSX.Location = new System.Drawing.Point(8, 112);
			this.lblSX.Name = "lblSX";
			this.lblSX.Size = new System.Drawing.Size(40, 16);
			this.lblSX.TabIndex = 15;
			this.lblSX.Text = "Size X:";
			// 
			// txtYSize
			// 
			this.txtYSize.Location = new System.Drawing.Point(104, 104);
			this.txtYSize.MaxLength = 4;
			this.txtYSize.Name = "txtYSize";
			this.txtYSize.Size = new System.Drawing.Size(40, 20);
			this.txtYSize.TabIndex = 14;
			this.txtYSize.Text = "500";
			// 
			// txtXSize
			// 
			this.txtXSize.Location = new System.Drawing.Point(48, 104);
			this.txtXSize.MaxLength = 4;
			this.txtXSize.Name = "txtXSize";
			this.txtXSize.Size = new System.Drawing.Size(40, 20);
			this.txtXSize.TabIndex = 13;
			this.txtXSize.Text = "1200";
			// 
			// lblBRK2
			// 
			this.lblBRK2.Enabled = false;
			this.lblBRK2.Location = new System.Drawing.Point(120, 64);
			this.lblBRK2.Name = "lblBRK2";
			this.lblBRK2.Size = new System.Drawing.Size(32, 16);
			this.lblBRK2.TabIndex = 12;
			this.lblBRK2.Text = "days.";
			// 
			// lblBRK1
			// 
			this.lblBRK1.Enabled = false;
			this.lblBRK1.Location = new System.Drawing.Point(8, 64);
			this.lblBRK1.Name = "lblBRK1";
			this.lblBRK1.Size = new System.Drawing.Size(64, 16);
			this.lblBRK1.TabIndex = 11;
			this.lblBRK1.Text = "Break every";
			// 
			// nmBreakpt
			// 
			this.nmBreakpt.Enabled = false;
			this.nmBreakpt.Location = new System.Drawing.Point(72, 64);
			this.nmBreakpt.Minimum = new System.Decimal(new int[] {
																	  1,
																	  0,
																	  0,
																	  0});
			this.nmBreakpt.Name = "nmBreakpt";
			this.nmBreakpt.Size = new System.Drawing.Size(48, 20);
			this.nmBreakpt.TabIndex = 10;
			this.nmBreakpt.Value = new System.Decimal(new int[] {
																	1,
																	0,
																	0,
																	0});
			this.nmBreakpt.ValueChanged += new System.EventHandler(this.nmBreakpt_ValueChanged);
			// 
			// lblSY
			// 
			this.lblSY.Location = new System.Drawing.Point(88, 112);
			this.lblSY.Name = "lblSY";
			this.lblSY.Size = new System.Drawing.Size(16, 16);
			this.lblSY.TabIndex = 15;
			this.lblSY.Text = "Y:";
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.lblLines);
			this.groupBox1.Controls.Add(this.chkLines);
			this.groupBox1.Controls.Add(this.lblPoints);
			this.groupBox1.Controls.Add(this.chkPts);
			this.groupBox1.Controls.Add(this.lblINC);
			this.groupBox1.Controls.Add(this.lblOUT);
			this.groupBox1.Controls.Add(this.lblNET);
			this.groupBox1.Controls.Add(this.chkNet);
			this.groupBox1.Controls.Add(this.chkOut);
			this.groupBox1.Controls.Add(this.chkInc);
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Controls.Add(this.rdoChange);
			this.groupBox1.Controls.Add(this.lblTotal);
			this.groupBox1.Controls.Add(this.rdoTotal);
			this.groupBox1.Location = new System.Drawing.Point(208, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(160, 120);
			this.groupBox1.TabIndex = 14;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Graph";
			// 
			// lblLines
			// 
			this.lblLines.Location = new System.Drawing.Point(88, 96);
			this.lblLines.Name = "lblLines";
			this.lblLines.Size = new System.Drawing.Size(64, 16);
			this.lblLines.TabIndex = 13;
			this.lblLines.Text = "Trendlines";
			// 
			// chkLines
			// 
			this.chkLines.Checked = true;
			this.chkLines.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkLines.Location = new System.Drawing.Point(72, 96);
			this.chkLines.Name = "chkLines";
			this.chkLines.Size = new System.Drawing.Size(40, 16);
			this.chkLines.TabIndex = 12;
			this.chkLines.Text = "checkBox1";
			// 
			// lblPoints
			// 
			this.lblPoints.Location = new System.Drawing.Point(24, 96);
			this.lblPoints.Name = "lblPoints";
			this.lblPoints.Size = new System.Drawing.Size(40, 16);
			this.lblPoints.TabIndex = 11;
			this.lblPoints.Text = "Points";
			// 
			// chkPts
			// 
			this.chkPts.Checked = true;
			this.chkPts.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkPts.Location = new System.Drawing.Point(8, 96);
			this.chkPts.Name = "chkPts";
			this.chkPts.Size = new System.Drawing.Size(16, 16);
			this.chkPts.TabIndex = 10;
			this.chkPts.Text = "checkBox1";
			// 
			// lblINC
			// 
			this.lblINC.Enabled = false;
			this.lblINC.Location = new System.Drawing.Point(96, 24);
			this.lblINC.Name = "lblINC";
			this.lblINC.Size = new System.Drawing.Size(56, 16);
			this.lblINC.TabIndex = 9;
			this.lblINC.Text = "Incoming";
			// 
			// lblOUT
			// 
			this.lblOUT.Enabled = false;
			this.lblOUT.Location = new System.Drawing.Point(96, 48);
			this.lblOUT.Name = "lblOUT";
			this.lblOUT.Size = new System.Drawing.Size(56, 16);
			this.lblOUT.TabIndex = 8;
			this.lblOUT.Text = "Outgoing";
			// 
			// lblNET
			// 
			this.lblNET.Enabled = false;
			this.lblNET.Location = new System.Drawing.Point(96, 72);
			this.lblNET.Name = "lblNET";
			this.lblNET.Size = new System.Drawing.Size(56, 16);
			this.lblNET.TabIndex = 7;
			this.lblNET.Text = "Net";
			// 
			// chkNet
			// 
			this.chkNet.Checked = true;
			this.chkNet.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkNet.Enabled = false;
			this.chkNet.Location = new System.Drawing.Point(80, 72);
			this.chkNet.Name = "chkNet";
			this.chkNet.Size = new System.Drawing.Size(16, 16);
			this.chkNet.TabIndex = 6;
			// 
			// chkOut
			// 
			this.chkOut.Enabled = false;
			this.chkOut.Location = new System.Drawing.Point(80, 48);
			this.chkOut.Name = "chkOut";
			this.chkOut.Size = new System.Drawing.Size(16, 16);
			this.chkOut.TabIndex = 5;
			// 
			// chkInc
			// 
			this.chkInc.Enabled = false;
			this.chkInc.Location = new System.Drawing.Point(80, 24);
			this.chkInc.Name = "chkInc";
			this.chkInc.Size = new System.Drawing.Size(16, 16);
			this.chkInc.TabIndex = 4;
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(24, 64);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(56, 16);
			this.label2.TabIndex = 3;
			this.label2.Text = "Changes";
			// 
			// rdoChange
			// 
			this.rdoChange.Location = new System.Drawing.Point(8, 64);
			this.rdoChange.Name = "rdoChange";
			this.rdoChange.Size = new System.Drawing.Size(16, 16);
			this.rdoChange.TabIndex = 2;
			this.rdoChange.CheckedChanged += new System.EventHandler(this.rdoChange_CheckedChanged);
			// 
			// lblTotal
			// 
			this.lblTotal.Location = new System.Drawing.Point(24, 24);
			this.lblTotal.Name = "lblTotal";
			this.lblTotal.Size = new System.Drawing.Size(40, 16);
			this.lblTotal.TabIndex = 1;
			this.lblTotal.Text = "Totals";
			// 
			// rdoTotal
			// 
			this.rdoTotal.Checked = true;
			this.rdoTotal.Location = new System.Drawing.Point(8, 24);
			this.rdoTotal.Name = "rdoTotal";
			this.rdoTotal.Size = new System.Drawing.Size(16, 16);
			this.rdoTotal.TabIndex = 0;
			this.rdoTotal.TabStop = true;
			this.rdoTotal.CheckedChanged += new System.EventHandler(this.rdoTotal_CheckedChanged);
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.label5);
			this.groupBox2.Controls.Add(this.label4);
			this.groupBox2.Controls.Add(this.label3);
			this.groupBox2.Controls.Add(this.rdoMonth);
			this.groupBox2.Controls.Add(this.rdoDay);
			this.groupBox2.Controls.Add(this.rdoWeek);
			this.groupBox2.Location = new System.Drawing.Point(192, 200);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(176, 48);
			this.groupBox2.TabIndex = 15;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Trendines by";
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(128, 24);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(40, 16);
			this.label5.TabIndex = 5;
			this.label5.Text = "Month";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(80, 24);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(24, 16);
			this.label4.TabIndex = 4;
			this.label4.Text = "Day";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(24, 24);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(40, 16);
			this.label3.TabIndex = 3;
			this.label3.Text = "Week";
			// 
			// rdoMonth
			// 
			this.rdoMonth.Location = new System.Drawing.Point(112, 24);
			this.rdoMonth.Name = "rdoMonth";
			this.rdoMonth.Size = new System.Drawing.Size(16, 16);
			this.rdoMonth.TabIndex = 2;
			this.rdoMonth.Text = "radioButton3";
			// 
			// rdoDay
			// 
			this.rdoDay.Location = new System.Drawing.Point(64, 24);
			this.rdoDay.Name = "rdoDay";
			this.rdoDay.Size = new System.Drawing.Size(16, 16);
			this.rdoDay.TabIndex = 1;
			this.rdoDay.Text = "radioButton2";
			// 
			// rdoWeek
			// 
			this.rdoWeek.Checked = true;
			this.rdoWeek.Location = new System.Drawing.Point(8, 24);
			this.rdoWeek.Name = "rdoWeek";
			this.rdoWeek.Size = new System.Drawing.Size(16, 16);
			this.rdoWeek.TabIndex = 0;
			this.rdoWeek.TabStop = true;
			// 
			// grpRet
			// 
			this.grpRet.Controls.Add(this.lblRetExc);
			this.grpRet.Controls.Add(this.lblRetInc);
			this.grpRet.Controls.Add(this.rdoRetExc);
			this.grpRet.Controls.Add(this.rdoRetInc);
			this.grpRet.Location = new System.Drawing.Point(16, 232);
			this.grpRet.Name = "grpRet";
			this.grpRet.Size = new System.Drawing.Size(120, 64);
			this.grpRet.TabIndex = 16;
			this.grpRet.TabStop = false;
			this.grpRet.Text = "Retired Members:";
			// 
			// lblRetExc
			// 
			this.lblRetExc.Location = new System.Drawing.Point(32, 40);
			this.lblRetExc.Name = "lblRetExc";
			this.lblRetExc.Size = new System.Drawing.Size(56, 16);
			this.lblRetExc.TabIndex = 3;
			this.lblRetExc.Text = "Exclude";
			// 
			// lblRetInc
			// 
			this.lblRetInc.Location = new System.Drawing.Point(32, 16);
			this.lblRetInc.Name = "lblRetInc";
			this.lblRetInc.Size = new System.Drawing.Size(56, 16);
			this.lblRetInc.TabIndex = 2;
			this.lblRetInc.Text = "Include";
			// 
			// rdoRetExc
			// 
			this.rdoRetExc.Location = new System.Drawing.Point(8, 40);
			this.rdoRetExc.Name = "rdoRetExc";
			this.rdoRetExc.Size = new System.Drawing.Size(16, 16);
			this.rdoRetExc.TabIndex = 1;
			this.rdoRetExc.Text = "radioButton1";
			// 
			// rdoRetInc
			// 
			this.rdoRetInc.Checked = true;
			this.rdoRetInc.Location = new System.Drawing.Point(8, 16);
			this.rdoRetInc.Name = "rdoRetInc";
			this.rdoRetInc.Size = new System.Drawing.Size(16, 16);
			this.rdoRetInc.TabIndex = 0;
			this.rdoRetInc.TabStop = true;
			this.rdoRetInc.Text = "radioButton1";
			// 
			// frmDKPOT
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(376, 319);
			this.Controls.Add(this.grpRet);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.grpOne);
			this.Controls.Add(this.lblEnd);
			this.Controls.Add(this.lblStart);
			this.Controls.Add(this.dateTimePicker2);
			this.Controls.Add(this.dateTimePicker1);
			this.Controls.Add(this.btnGenerate);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.Name = "frmDKPOT";
			this.Text = "DKP Over Time Grapher";
			this.grpOne.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.nmBreakpt)).EndInit();
			this.groupBox1.ResumeLayout(false);
			this.groupBox2.ResumeLayout(false);
			this.grpRet.ResumeLayout(false);
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

		#region Main Graphing Function
		private void Graph(DateTime start, DateTime end)
		{
			if (start.CompareTo(end)>=0) return;
			int width = int.Parse(txtXSize.Text);
			int height = int.Parse(txtYSize.Text);
			int xoffset = 75; // offsets for axes/labels
			int yoffset = 40;
			int lines = 1; // # of subsets of points
			int points; // total # of points
			double xscalar;
			double yscalar;
			double total;
			int curline = 0;
			int curpoint = 0;
			Point[][] pa;
			Point[][] paInc;
			Point[][] paOut;
		
			DataTable dbCount = null; // Count of days
			DataTable dbDailyNet = null; // daily net change
			DataTable dbDailyInc = null; // daily incoming
			DataTable dbDailyOut = null; // daily outgoing
			DataTable dbTotal = null; // total up to start - 1

			string query;
			string filename = parent.OutputDirectory;
			filename += start.ToString("yyyymmdd") + end.ToString("yyyyddmm");
			if (rdoTotal.Checked) filename += "t";
			else 
			{
				filename += "c";
				if (chkNet.Checked) filename += "n";
				if (chkInc.Checked) filename += "i";
				if (chkOut.Checked) filename += "o";
			}
			if (chkPts.Checked) filename += "p";
			if (chkLines.Checked) 
			{
				filename += "l";
				if (rdoDay.Checked) filename += "d";
				else if (rdoWeek.Checked) filename += "w";
				else if (rdoMonth.Checked) filename += "m";
			}
			if (rdoRetExc.Checked) filename += "e";
			filename += width + "x" + height;
			filename += ".bmp";
			OleDbConnection dbConnect = null;
			OleDbDataAdapter dkpDA;
			retired = rdoRetInc.Checked;
			DateTime currentWeek = start;
			while (rdoWeek.Checked&&currentWeek.DayOfWeek != DayOfWeek.Sunday) currentWeek = currentWeek.AddDays(-1);
			string connectionStr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + parent.DBString;
			try { dbConnect = new OleDbConnection(connectionStr); }
			catch (Exception ex) 
			{
				MessageBox.Show("Could not create data connection. \n(" + ex.Message + ")","Error");
			}

			TimeSpan t = end - start;
			xscalar = ((double)width-xoffset)/t.Days;
			try
			{
				dbConnect.Open();
				dbCount = new DataTable();
				dbDailyNet = new DataTable();
				dbDailyInc = new DataTable();
				dbDailyOut = new DataTable();
				dbTotal = new DataTable();
				query = "SELECT Count(Date) FROM (SELECT DISTINCT Date FROM DKS WHERE DKS.Date >=#" + start.Month + "/" + start.Day + "/" + start.Year + "# AND DKS.Date<=#" + end.Month + "/" + end.Day + "/" + end.Year + "#)";
				dkpDA = new OleDbDataAdapter(query,dbConnect);
				dkpDA.Fill(dbCount);
				query = "SELECT Sum(PTS) AS SumDKP FROM DKS WHERE DKS.Date <#" + start.Month + "/" + start.Day + "/" + start.Year + "# ";
				if (!retired) query += " AND (DKS.Name<>'zzzDKP Retired')";
				else query += "AND (DKS.Comment<>'Mr Epic' AND DKS.Comment<>'Mr. App-Alt' AND DKS.Comment <> 'Mr. Progress' AND DKS.Comment <> 'mr. progression')";
				dkpDA = new OleDbDataAdapter(query,dbConnect);
				dkpDA.Fill(dbTotal);
				query = "SELECT Sum(PTS) AS SumDKP,Date FROM DKS WHERE DKS.Date >=#" + start.Month + "/" + start.Day + "/" + start.Year + "# AND DKS.Date<=#" + end.Month + "/" + end.Day + "/" + end.Year + "# ";
				if (!retired) query += " AND (DKS.Name<>'zzzDKP Retired')";
				else query += "AND (DKS.Comment<>'Mr Epic' AND DKS.Comment<>'Mr. App-Alt' AND DKS.Comment <> 'Mr. Progress' AND DKS.Comment <> 'mr. progression')";
				query += " GROUP BY Date ORDER BY Date ASC";
				dkpDA = new OleDbDataAdapter(query,dbConnect);
				dkpDA.Fill(dbDailyNet);
				query = "SELECT Sum(PTS) AS SumDKP,Date FROM DKS WHERE DKS.Date >=#" + start.Month + "/" + start.Day + "/" + start.Year + "# AND DKS.Date<=#" + end.Month + "/" + end.Day + "/" + end.Year + "# AND PTS>=0 ";
				if (!retired) query += " AND (DKS.Name<>'zzzDKP Retired')";
				else query += "AND (DKS.Comment<>'Mr Epic' AND DKS.Comment<>'Mr. App-Alt' AND DKS.Comment <> 'Mr. Progress' AND DKS.Comment <> 'mr. progression')";
				query += " GROUP BY Date ORDER BY Date ASC";
				dkpDA = new OleDbDataAdapter(query,dbConnect);
				dkpDA.Fill(dbDailyInc);
				query = "SELECT Sum(PTS) AS SumDKP,Date FROM DKS WHERE DKS.Date >=#" + start.Month + "/" + start.Day + "/" + start.Year + "# AND DKS.Date<=#" + end.Month + "/" + end.Day + "/" + end.Year + "# AND PTS<0 ";
				if (!retired) query += " AND (DKS.Name<>'zzzDKP Retired')";
				else query += "AND (DKS.Comment<>'Mr Epic' AND DKS.Comment<>'Mr. App-Alt' AND DKS.Comment <> 'Mr. Progress' AND DKS.Comment <> 'mr. progression')";
				query += " GROUP BY Date ORDER BY Date ASC";
				dkpDA = new OleDbDataAdapter(query,dbConnect);
				dkpDA.Fill(dbDailyOut);
			}
			catch (Exception ex) 
			{
				MessageBox.Show("Could not open data connection. \n(" + ex.Message + ")","Error");
			}
			finally { dbConnect.Close(); }
			points = (int)dbCount.Rows[0].ItemArray[0];
			total = (double)dbTotal.Rows[0].ItemArray[0];
			if (rdoWeek.Checked) lines = ((int)t.TotalDays/7)+2;
			else if (rdoDay.Checked) lines = (1);
			else if (rdoMonth.Checked) lines = ((int)t.TotalDays/28)+2;
			pa = new Point[lines][];
			paInc = new Point[lines][];
			paOut = new Point[lines][];
			for (int i=0;i<lines;i++) 
			{
				if (rdoWeek.Checked) 
				{
					pa[i] = new Point[7];
					paInc[i] = new Point[7];
					paOut[i] = new Point[7];
				}
				else if (rdoMonth.Checked) 
				{
					pa[i] = new Point[31];
					paInc[i] = new Point[31];
					paOut[i] = new Point[31];
				}
				else if (rdoDay.Checked) 
				{
					pa[i] = new Point[points+1];
					paInc[i] = new Point[points+1];
					paOut[i] = new Point[points+1];
					break;
				}
			}

			foreach (DataRow r in dbDailyNet.Rows)
			{	
				DateTime d = (DateTime)r["Date"];
				TimeSpan u = d - currentWeek;
				TimeSpan v = d - start;
				if (rdoWeek.Checked&&u.Days>=7) { curline++; curpoint = 0; currentWeek = currentWeek.AddDays(7); }
				else if (rdoDay.Checked) { }
				else if (rdoMonth.Checked&&(d.Month-currentWeek.Month!=0)) { curline++; curpoint = 0; currentWeek = currentWeek.AddMonths(1); }
				if (!rdoTotal.Checked) pa[curline][curpoint] = new Point((int)(xscalar*v.Days+xoffset),(int)(double)r["SumDKP"]);
				else { total += (double)r["SumDKP"]; pa[curline][curpoint] = new Point((int)(xscalar*v.Days+xoffset),(int)total); }
				curpoint++;
			}
			currentWeek = start;
			while (rdoWeek.Checked&&currentWeek.DayOfWeek != DayOfWeek.Sunday) currentWeek = currentWeek.AddDays(-1);
			curpoint = curline = 0;
			foreach (DataRow r in dbDailyInc.Rows)
			{	
				DateTime d = (DateTime)r["Date"];
				TimeSpan u = d - currentWeek;
				TimeSpan v = d - start;
				if (rdoWeek.Checked&&u.Days>=7) { curline++; curpoint = 0; currentWeek = currentWeek.AddDays(7); }
				else if (rdoDay.Checked) { }
				else if (rdoMonth.Checked&&(d.Month-currentWeek.Month!=0)) { curline++; curpoint = 0; currentWeek = currentWeek.AddMonths(1); }
				if (!rdoTotal.Checked) paInc[curline][curpoint] = new Point((int)(xscalar*v.Days+xoffset),(int)(double)r["SumDKP"]);
				curpoint++;
			}
			curpoint = curline = 0;
			currentWeek = start;
			while (rdoWeek.Checked&&currentWeek.DayOfWeek != DayOfWeek.Sunday) currentWeek = currentWeek.AddDays(-1);
			foreach (DataRow r in dbDailyOut.Rows)
			{	
				DateTime d = (DateTime)r["Date"];
				TimeSpan u = d - currentWeek;
				TimeSpan v = d - start;
				if (rdoWeek.Checked&&u.Days>=7) { curline++; curpoint = 0; currentWeek = currentWeek.AddDays(7); }
				else if (rdoDay.Checked) {  }
				else if (rdoMonth.Checked&&(d.Month-currentWeek.Month!=0)) { curline++; curpoint = 0; currentWeek = currentWeek.AddMonths(1); }
				if (!rdoTotal.Checked) paOut[curline][curpoint] = new Point((int)(xscalar*v.Days+xoffset),(int)(double)r["SumDKP"]);
				curpoint++;
			}
			Point extrema = Line.maxmin(pa);
			if (!rdoTotal.Checked&&chkInc.Checked) 
			{
				Point w = Line.maxmin(paInc);
				if (w.X<extrema.X) extrema.X = w.X;
				if (w.Y>extrema.Y) extrema.Y = w.Y;
			}
			if (!rdoTotal.Checked&&chkOut.Checked) 
			{
				Point w = Line.maxmin(paOut);
				if (w.X<extrema.X) extrema.X = w.X;
				if (w.Y>extrema.Y) extrema.Y = w.Y;
			}
			yscalar = (double)(height-yoffset-5)/(extrema.Y - extrema.X);
			
					
			Brush br = new SolidBrush(bgColor);
			Pen p = new Pen(axisColor,axisWidth);
			Bitmap bit = new Bitmap(width,height);
			Graphics g = Graphics.FromImage(bit);
			g.SmoothingMode = SmoothingMode.AntiAlias;

			// set background
			g.FillRectangle(br,0,0,width,height);

			// axes
			g.DrawLine(p,new Point(xoffset,height-yoffset),new Point(xoffset,0));
			g.DrawLine(p,new Point(xoffset,height-yoffset),new Point(width,height-yoffset));

			// labels
			br = new SolidBrush(textColor);
			string str = "Date";
			RectangleF label = new RectangleF(width/2-(str.Length/2)*labelFont.SizeInPoints,height-labelFont.GetHeight(g),str.Length*labelFont.SizeInPoints,labelFont.GetHeight(g));
			g.DrawString(str,labelFont,br,label);
			str = "Value";
			label = new RectangleF(0,(height-yoffset)/2-(str.Length/2)*labelFont.GetHeight(g),labelFont.SizeInPoints,labelFont.GetHeight(g)*str.Length);
			g.DrawString(str,labelFont,br,label);


			StringFormat sf = new StringFormat();
			sf.Alignment = StringAlignment.Center;
			str = start.ToString("ddMMMyyyy");
			label = new RectangleF(xoffset-xlabelspace/2,height-yoffset+crossLength,xlabelspace,labelFontSmall.GetHeight(g));
			g.DrawString(str,labelFontSmall,br,label,sf);
			DateTime cur = start;
			int years = 0;
			if (t.Days>=width/xlabelspace+1) 
			{ 
				double ts = t.Days/((double)width/xlabelspace);
				for (int xpos=xlabelspace+xoffset;xpos<=width;xpos+=xlabelspace)
				{
					g.DrawLine(p,xpos,height-yoffset-crossLength,xpos,height-yoffset+crossLength);
					cur = cur.AddDays(ts);
					if (cur.Year - start.Year>years) { str = cur.ToString("ddMMMyyyy"); years = cur.Year - start.Year; }
					else str = cur.ToString("ddMMM");
					label = new RectangleF(xpos-xlabelspace/2,height-yoffset+crossLength,xlabelspace,labelFontSmall.GetHeight(g));
					g.DrawString(str,labelFontSmall,br,label,sf);
				} 
			} 
			else
			{
				for (int i = 1; i<= t.Days; i++)
				{
					cur = cur.AddDays(1);
					int xpos = (int) ( ((double)width)/t.Days * i + xoffset );
					g.DrawLine(p,xpos,height-yoffset-crossLength,xpos,height-yoffset+crossLength);
					if (cur.Year - start.Year>years) { str = cur.ToString("ddMMMyyyy"); years = cur.Year - start.Year; }
					else str = cur.ToString("ddMMM");
					label = new RectangleF(xpos-xlabelspace/2,height-yoffset+crossLength,xlabelspace,labelFontSmall.GetHeight(g));
					g.DrawString(str,labelFontSmall,br,label,sf);
				}
			}

			sf.Alignment = StringAlignment.Far;
			str = extrema.X.ToString();
			label = new RectangleF(labelFont.SizeInPoints,height-yoffset-labelFontSmall.GetHeight(g),xoffset-labelFont.SizeInPoints-crossLength,labelFontSmall.GetHeight(g));
			g.DrawString(str,labelFontSmall,br,label,sf);
			for (int ypos=height-yoffset-ylabelspace; ypos>=0; ypos-=ylabelspace)
			{
				double ym = ((double)extrema.Y-extrema.X)/(height-yoffset);
				g.DrawLine(p,xoffset-crossLength,ypos,xoffset+crossLength,ypos); //height-yoffset-ypos
				int thisy = (int)(ym*(height-yoffset-ypos)+extrema.X);
				str = thisy.ToString();
				label = new RectangleF(labelFont.SizeInPoints,ypos-labelFontSmall.GetHeight(g)/2,xoffset-labelFont.SizeInPoints-crossLength,labelFontSmall.GetHeight(g));
				g.DrawString(str,labelFontSmall,br,label,sf);
			}
            		

			Point lastLine = new Point(xoffset,height-yoffset);
			if (rdoTotal.Checked||chkNet.Checked) 
			{
				if (rdoDay.Checked) 
				{
					foreach (Point b in pa[0])
					{
						Point c = b;
						if (c.IsEmpty) continue;
						c.Y = (int)(height-5-c.Y*yscalar+extrema.X*yscalar-yoffset);
						p.Color = crossNetColor;
						p.Width = pointWidth;
						if (chkPts.Checked) g.DrawLine(p,new Point(c.X-crossLength,c.Y),new Point(c.X+crossLength,c.Y));
						if (chkPts.Checked) g.DrawLine(p,new Point(c.X,c.Y-crossLength),new Point(c.X,c.Y+crossLength));
						p.Color = lineNetColor;
						p.Width = lineWidth;
						if (lastLine.X!=xoffset&&chkLines.Checked) g.DrawLine(p,lastLine,c);
						lastLine = new Point(c.X,c.Y);
					}
				}
				else 
				{
					foreach (Point[] a in pa) 
					{
						bool allempty = true;
						Point s = new Point(width+1,0);
						Point e = new Point(0,0);
						p.Color = crossNetColor;
						p.Width = pointWidth;
						foreach(Point b in a)
						{
							if (b.IsEmpty) continue;
							allempty = false;
							if (b.X<s.X) { s.X = b.X; s.Y = b.Y; }
							if (b.X>e.X) { e.X = b.X; e.Y = b.Y; }
							Point c = b;
							c.Y = (int)(height-5-c.Y*yscalar+extrema.X*yscalar-yoffset);
							if (chkPts.Checked) g.DrawLine(p,new Point(c.X-crossLength,c.Y),new Point(c.X+crossLength,c.Y));
							if (chkPts.Checked) g.DrawLine(p,new Point(c.X,c.Y-crossLength),new Point(c.X,c.Y+crossLength));
						}
						if (allempty) continue;
						
						p.Color = lineNetColor;
						p.Width = lineWidth;
						Line l = Line.BestFit(a);
						if (lastLine.X!=xoffset&&!s.Equals(e)&&chkLines.Checked) g.DrawLine(p,lastLine,new Point(s.X,(int)(height-5-l.eval(s.X)*yscalar+extrema.X*yscalar-yoffset)));
						if (!s.Equals(e)&&chkLines.Checked) g.DrawLine(p,new Point(s.X,(int)(height-5-l.eval(s.X)*yscalar+extrema.X*yscalar-yoffset)),new Point(e.X,(int)(height-5-l.eval(e.X)*yscalar+extrema.X*yscalar-yoffset)));
						if (s.Equals(e)) 
						{
							if(chkLines.Checked) g.DrawLine(p,lastLine,new Point(s.X,(int)(height-5-s.Y*yscalar+extrema.X*yscalar-yoffset)));
							lastLine = new Point(s.X,(int)(height-5-s.Y*yscalar+extrema.X*yscalar-yoffset));
						}
						if (!s.Equals(e)) lastLine = new Point(e.X,(int)(height-5-l.eval(e.X)*yscalar+extrema.X*yscalar-yoffset));
					}
				}
			}

			lastLine = new Point(xoffset,height-yoffset);

			if (rdoChange.Checked&&chkInc.Checked)
			{
				if (rdoDay.Checked) 
				{
					foreach (Point b in paInc[0])
					{
						Point c = b;
						if (c.IsEmpty) continue;
						c.Y = (int)(height-5-c.Y*yscalar+extrema.X*yscalar-yoffset);
						p.Color = crossIncColor;
						p.Width = pointWidth;
						if (chkPts.Checked) g.DrawLine(p,new Point(c.X-crossLength,c.Y),new Point(c.X+crossLength,c.Y));
						if (chkPts.Checked) g.DrawLine(p,new Point(c.X,c.Y-crossLength),new Point(c.X,c.Y+crossLength));
						p.Color = lineIncColor;
						p.Width = lineWidth;
						if (lastLine.X!=xoffset&&chkLines.Checked) g.DrawLine(p,lastLine,c);
						lastLine = new Point(c.X,c.Y);
					}
				}
				else 
				{
					foreach (Point[] a in paInc) 
					{
						bool allempty = true;
						Point s = new Point(width+1,0);
						Point e = new Point(0,0);
						p.Color = crossIncColor;
						p.Width = pointWidth;
						foreach(Point b in a)
						{
							if (b.IsEmpty) continue;
							allempty = false;
							if (b.X<s.X) { s.X = b.X; s.Y = b.Y; }
							if (b.X>e.X) { e.X = b.X; e.Y = b.Y; }
							Point c = b;
							c.Y = (int)(height-5-c.Y*yscalar+extrema.X*yscalar-yoffset);
							if (chkPts.Checked) g.DrawLine(p,new Point(c.X-crossLength,c.Y),new Point(c.X+crossLength,c.Y));
							if (chkPts.Checked) g.DrawLine(p,new Point(c.X,c.Y-crossLength),new Point(c.X,c.Y+crossLength));
						}
						if (allempty) continue;
						
						p.Color = lineIncColor;
						p.Width = lineWidth;
						Line l = Line.BestFit(a);
						if (lastLine.X!=xoffset&&!s.Equals(e)&&chkLines.Checked) g.DrawLine(p,lastLine,new Point(s.X,(int)(height-5-l.eval(s.X)*yscalar+extrema.X*yscalar-yoffset)));
						if (!s.Equals(e)&&chkLines.Checked) g.DrawLine(p,new Point(s.X,(int)(height-5-l.eval(s.X)*yscalar+extrema.X*yscalar-yoffset)),new Point(e.X,(int)(height-5-l.eval(e.X)*yscalar+extrema.X*yscalar-yoffset)));
						if (s.Equals(e)) 
						{
							if (chkLines.Checked) g.DrawLine(p,lastLine,new Point(s.X,(int)(height-5-s.Y*yscalar+extrema.X*yscalar-yoffset)));
							lastLine = new Point(s.X,(int)(height-5-s.Y*yscalar+extrema.X*yscalar-yoffset));
						}
						if (!s.Equals(e)) lastLine = new Point(e.X,(int)(height-5-l.eval(e.X)*yscalar+extrema.X*yscalar-yoffset));
					}
				}
			}

			lastLine = new Point(xoffset,height-yoffset);

			if (rdoChange.Checked&&chkOut.Checked)
			{
				if (rdoDay.Checked) 
				{
					foreach (Point b in paOut[0])
					{
						Point c = b;
						if (c.IsEmpty) continue;
						c.Y = (int)(height-5-c.Y*yscalar+extrema.X*yscalar-yoffset);
						p.Color = crossOutColor;
						p.Width = pointWidth;
						if (chkPts.Checked) g.DrawLine(p,new Point(c.X-crossLength,c.Y),new Point(c.X+crossLength,c.Y));
						if (chkPts.Checked) g.DrawLine(p,new Point(c.X,c.Y-crossLength),new Point(c.X,c.Y+crossLength));
						p.Color = lineOutColor;
						p.Width = lineWidth;
						if (lastLine.X!=xoffset&&chkLines.Checked) g.DrawLine(p,lastLine,c);
						lastLine = new Point(c.X,c.Y);
					}
				}
				else 
				{
					foreach (Point[] a in paOut) 
					{
						bool allempty = true;
						Point s = new Point(width+1,0);
						Point e = new Point(0,0);
						p.Color = crossOutColor;
						p.Width = pointWidth;
						foreach(Point b in a)
						{
							if (b.IsEmpty) continue;
							allempty = false;
							if (b.X<s.X) { s.X = b.X; s.Y = b.Y; }
							if (b.X>e.X) { e.X = b.X; e.Y = b.Y; }
							Point c = b;
							c.Y = (int)(height-5-c.Y*yscalar+extrema.X*yscalar-yoffset);
							if (chkPts.Checked) g.DrawLine(p,new Point(c.X-crossLength,c.Y),new Point(c.X+crossLength,c.Y));
							if (chkPts.Checked) g.DrawLine(p,new Point(c.X,c.Y-crossLength),new Point(c.X,c.Y+crossLength));
						}
						if (allempty) continue;
						
						p.Color = lineOutColor;
						p.Width = lineWidth;
						Line l = Line.BestFit(a);
						if (lastLine.X!=xoffset&&!s.Equals(e)&&chkLines.Checked) g.DrawLine(p,lastLine,new Point(s.X,(int)(height-5-l.eval(s.X)*yscalar+extrema.X*yscalar-yoffset)));
						if (!s.Equals(e)&&chkLines.Checked) g.DrawLine(p,new Point(s.X,(int)(height-5-l.eval(s.X)*yscalar+extrema.X*yscalar-yoffset)),new Point(e.X,(int)(height-5-l.eval(e.X)*yscalar+extrema.X*yscalar-yoffset)));
						if (s.Equals(e)) 
						{
							if (chkLines.Checked) g.DrawLine(p,lastLine,new Point(s.X,(int)(height-5-s.Y*yscalar+extrema.X*yscalar-yoffset)));
							lastLine = new Point(s.X,(int)(height-5-s.Y*yscalar+extrema.X*yscalar-yoffset));
						}
						if (!s.Equals(e)) lastLine = new Point(e.X,(int)(height-5-l.eval(e.X)*yscalar+extrema.X*yscalar-yoffset));
					}
				}
			}

			bit.Save(filename,ImageFormat.Bmp);
		}
		#endregion

		#region Events
		private void btnGenerate_Click(object sender, System.EventArgs e)
		{
			breakpt = (int)nmBreakpt.Value;
			if (rdoSingle.Checked) Graph(dateTimePicker1.Value,dateTimePicker2.Value);
			else 
			{
				TimeSpan t = dateTimePicker2.Value - dateTimePicker1.Value;
				if (t.Days<=breakpt) 
				{
					Graph(dateTimePicker1.Value,dateTimePicker2.Value);
				}
				else 
				{ 
					DateTime cur = dateTimePicker1.Value;
					TimeSpan u =  dateTimePicker2.Value - cur;
					while (u.Days>breakpt)
					{
						DateTime nex = cur.AddDays(breakpt);
						Console.WriteLine("Graphing: " + cur + " to " + nex + "(breakpt: " + breakpt + ")");
						Graph(cur,nex);
						cur = cur.AddDays(breakpt);
						u = dateTimePicker2.Value - cur;
					}
					Graph(cur,dateTimePicker2.Value);
				}
				
			}
		}
		private void rdoMulti_CheckedChanged(object sender, System.EventArgs e)
		{
			if (rdoMulti.Checked) 
			{
				breakpt = (int)nmBreakpt.Value;
				lblBRK1.Enabled = true;
				lblBRK2.Enabled = true;
				nmBreakpt.Enabled = true;
			}
			else 
			{
				breakpt = 0;
				lblBRK1.Enabled = false;
				lblBRK2.Enabled = false;
				nmBreakpt.Enabled = false;
			}

            		
		}
		private void rdoSingle_CheckedChanged(object sender, System.EventArgs e)
		{
			if (rdoMulti.Checked) 
			{
				breakpt = (int)nmBreakpt.Value;
				lblBRK1.Enabled = true;
				lblBRK2.Enabled = true;
				nmBreakpt.Enabled = true;
			}
			else 
			{
				breakpt = 0;
				lblBRK1.Enabled = false;
				lblBRK2.Enabled = false;
				nmBreakpt.Enabled = false;
			}
		
		}
		private void rdoTotal_CheckedChanged(object sender, System.EventArgs e)
		{
			if (rdoTotal.Checked) 
			{
				lblINC.Enabled = lblOUT.Enabled = lblNET.Enabled = chkInc.Enabled = chkOut.Enabled = chkNet.Enabled = false;
			}
			else
			{
				lblINC.Enabled = lblOUT.Enabled = lblNET.Enabled = chkInc.Enabled = chkOut.Enabled = chkNet.Enabled = true;
			}
		}
		private void rdoChange_CheckedChanged(object sender, System.EventArgs e)
		{
			if (rdoTotal.Checked) 
			{
				lblINC.Enabled = lblOUT.Enabled = lblNET.Enabled = chkInc.Enabled = chkOut.Enabled = chkNet.Enabled = false;
			}
			else
			{
				lblINC.Enabled = lblOUT.Enabled = lblNET.Enabled = chkInc.Enabled = chkOut.Enabled = chkNet.Enabled = true;
			}
		
		}
		private void nmBreakpt_ValueChanged(object sender, System.EventArgs e)
		{
			breakpt = (int)nmBreakpt.Value;
		}
		#endregion

	}
}
