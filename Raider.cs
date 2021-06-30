using System;

namespace ES_DKP_Utils
{
	public class Raider : IComparable
    {
        #region Constants
        public const double NODKP = -99999;
		public const string NOTIER = "E";
        public const double NOATTENDANCE = 0;
        #endregion

        #region Settings
        public double MINDKP = 0;
        #endregion

        #region Declarations
        public string Person { get; set; }
        public string Tier { get; set; }
        public System.Double DKP { get; set; }
        public System.Double AttendancePCT { get; set; }
        public string TellType { get; set; }
        public int Upgradeable { get; set; }
        public int Upgrades { get; set; }
		public string ItemMessage { get; set; }

        private frmMain owner;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
		#endregion

		#region Constructor
		public Raider(frmMain owner, string p, string t, double d, double a, string tellType = "N/A", int upgradeable = -1, int upgrades = -1, string itemMessage = "")
		{
            log.Debug($"Begin Method: Raider.Raider(frmMain,string,string,double,double,string,string) ({owner}, {p}, {t}, {d}, {a}, {tellType}, {upgradeable}, {upgrades}, {itemMessage}");

			this.owner = owner;
			Person = p;
			Tier = t;
			DKP = d;
            AttendancePCT = a;
            MINDKP = owner.MinDKP;
			TellType = tellType;
			ItemMessage = itemMessage;
            Upgradeable = upgradeable;
            Upgrades = upgrades;

            log.Debug("End Method: Raider.Raider()");
		}
        #endregion

        #region Methods
        public override string ToString()
		{
			return Person + ": " + DKP + " " + Tier + " (" + AttendancePCT + ")";
		}

		public int CompareTo(object obj)
		{
			if (obj is Raider) 
			{
				Raider that = (Raider)obj;

				if (owner.ItemDKP.Equals("A"))
				{
                    if (this.Tier.Equals("A") && (!that.Tier.Equals("A")) && (this.DKP > MINDKP)) return -1;
                    else if (that.Tier.Equals("A") && (!this.Tier.Equals("A")) && (that.DKP > MINDKP)) return 1;
                    else if (this.Tier.Equals("B") && (!that.Tier.Equals("A") && !that.Tier.Equals("B")) && (this.DKP > MINDKP)) return -1;
                    else if (that.Tier.Equals("B") && (!this.Tier.Equals("A") && !this.Tier.Equals("B")) && (that.DKP > MINDKP)) return 1;
                    else if (this.Tier.Equals("C") && (!that.Tier.Equals("A") && !that.Tier.Equals("B") && !that.Tier.Equals("C")) && (this.DKP > MINDKP)) return -1;
                    else if (that.Tier.Equals("C") && (!this.Tier.Equals("A") && !this.Tier.Equals("B") && !this.Tier.Equals("C")) && (that.DKP > MINDKP)) return 1;
                    else if (this.Tier.Equals(Raider.NOTIER)) return 1;
                    else if (that.Tier.Equals(Raider.NOTIER)) return -1;
                    else if (this.DKP < that.DKP) return 1;
                    else return -1;
				}

				else if (owner.ItemDKP.Equals("B"))
				{
                    if ((this.Tier.Equals("A") || this.Tier.Equals("B")) && (!that.Tier.Equals("A")) && (!that.Tier.Equals("B")) && (this.DKP > MINDKP)) return -1;
                    else if ((that.Tier.Equals("A") || that.Tier.Equals("B")) && (!this.Tier.Equals("A")) && (!this.Tier.Equals("B")) && (that.DKP > MINDKP)) return 1;
                    else if (this.Tier.Equals("C") && !that.Tier.Equals("A") && !that.Tier.Equals("B") && !that.Tier.Equals("C") && (this.DKP > MINDKP)) return -1;
                    else if (that.Tier.Equals("C") && !this.Tier.Equals("A") && !this.Tier.Equals("B") && !this.Tier.Equals("C") && (that.DKP > MINDKP)) return 1;
                    else if (this.Tier.Equals(Raider.NOTIER)) return 1;
                    else if (that.Tier.Equals(Raider.NOTIER)) return -1;
                    else if (this.DKP < that.DKP) return 1;
                    else return -1;
				}

				else if (owner.ItemDKP.Equals("C"))
				{
					if (this.Tier.Equals(Raider.NOTIER)) return 1;
					else if (that.Tier.Equals(Raider.NOTIER)) return -1;
					else if (this.DKP < that.DKP) return 1;
					else return -1;
				}

                else if (owner.ItemDKP.Equals("ATTD"))
                {
                    if (this.AttendancePCT > that.AttendancePCT) return -1;
                    else return 1;
                }

			}
			else throw new ArgumentException("Object to compare to is not a Raider object.");
			return 0;
        }
        #endregion
    }
}
