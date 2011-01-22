using System;
using System.Drawing;

namespace ES_DKP_Utils
{
	public class Line
	{
		private double slope;
		private double b;

		public Line(double m, Point p)
		{
			slope = m;
			b = p.Y - (p.X)* slope;
		}

		public Line(double m, double yi)
		{
			slope = m;
			b = yi;
		}

		public double eval(double x)
		{
			return slope*x+b;
		}

		public static Point maxmin(Point[][] p)
		{
			int max = int.MinValue;
			int min = int.MaxValue;
			foreach (Point[] q in p)
			{
				foreach (Point r in q)
				{
					if (r.Y>max&&!r.IsEmpty) max = r.Y;
					if (r.Y<min&&!r.IsEmpty) min = r.Y;
				}
			}
			Point o = new Point(min, max);
			return o;
		}

		public static Line BestFit(Point[] p)
		{
			int n = p.GetLength(0);
			double m = 0.0;
			double yi = 0.0;
			double sigxy = 0.0;
			double sigx = 0.0;
			double sigy = 0.0;
			double sigxx = 0.0;

			foreach (Point q in p)
			{
				if (q.IsEmpty) { n--; continue; }
				sigxy += q.X*q.Y;
				sigx += q.X;
				sigy += q.Y;
				sigxx += q.X*q.X;
			}

			m = (n*sigxy - sigx*sigy)/(n*sigxx-sigx*sigx);
			yi = (sigy-m*sigx)/n;

			Line l = new Line(m,yi);

			return l;

		}

		public override string ToString()
		{
			return "y = " + slope + "x " + (b<0?"- ":"+ ") + Math.Abs(b);
		}
	}
}
