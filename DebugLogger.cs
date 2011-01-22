using System;
using System.IO;
using System.Diagnostics;
using System.Collections;

namespace ES_DKP_Utils
{
	public class DebugLogger
	{
        private static string debugDirectory = @"G:\ES DKP Utils Debug\";
        private string filename;
        private ArrayList buffer;
        private System.Timers.Timer WriteLogTimer;

		public DebugLogger(string filename)
		{
            this.filename = filename;
            buffer = new ArrayList();
            WriteLogTimer = new System.Timers.Timer(15000);
            WriteLogTimer.Elapsed += new System.Timers.ElapsedEventHandler(WriteLogTimer_Elapsed);
            WriteLogTimer.Start();
		}

        public void WriteLogTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            StreamWriter debugger = null;
            try { debugger = new StreamWriter(debugDirectory + this.filename, true); }
            catch (Exception ex) {
                this.WriteLogTimer_Elapsed(null, null);
                return;
            }

            string[] k = new string[buffer.Count];
            buffer.CopyTo(k);
            buffer.Clear();
            foreach (string j in k)
            {
                debugger.WriteLine(j);
            }          

            debugger.Close();
        }

		[Conditional("DEBUG_1")]
		public void WriteDebug_1(string s)
		{			
            string t;
			t = "[" + DateTime.Now.ToString("dd/MM/yy HH:mm:ss.fff") + "] " + s;
            buffer.Add(t);
		}

		[Conditional("DEBUG_2")]
		public void WriteDebug_2(string s)
		{
			string t;
			t = "[" + DateTime.Now.ToString("dd/MM/yy HH:mm:ss.fff") + "] " + s;
            buffer.Add(t);
		}

		[Conditional("DEBUG_3")]
		public void WriteDebug_3(string s) 
        {
			string t;
			t = "[" + DateTime.Now.ToString("dd/MM/yy HH:mm:ss.fff") + "] " + s;
            buffer.Add(t);
		}
	}
}
