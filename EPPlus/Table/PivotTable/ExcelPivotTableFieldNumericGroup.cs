using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// A pivot table field numeric grouping.
	/// </summary>
	public class ExcelPivotTableFieldNumericGroup : ExcelPivotTableFieldGroup
	{
		#region Constants
		private const string StartPath = "d:fieldGroup/d:rangePr/@startNum";
		private const string EndPath = "d:fieldGroup/d:rangePr/@endNum";
		private const string GroupIntervalPath = "d:fieldGroup/d:rangePr/@groupInterval";
		#endregion

		#region Properties
		/// <summary>
		/// Gets the start value.
		/// </summary>
		public double Start
		{
			get
			{
				return (double)base.GetXmlNodeDoubleNull(ExcelPivotTableFieldNumericGroup.StartPath);
			}
			private set
			{
				base.SetXmlNodeString(ExcelPivotTableFieldNumericGroup.StartPath, value.ToString(CultureInfo.InvariantCulture));
			}
		}

		/// <summary>
		/// Gets the end value.
		/// </summary>
		public double End
		{
			get
			{
				return (double)base.GetXmlNodeDoubleNull(ExcelPivotTableFieldNumericGroup.EndPath);
			}
			private set
			{
				base.SetXmlNodeString(ExcelPivotTableFieldNumericGroup.EndPath, value.ToString(CultureInfo.InvariantCulture));
			}
		}

		/// <summary>
		/// Gets the interval.
		/// </summary>
		public double Interval
		{
			get
			{
				return (double)base.GetXmlNodeDoubleNull(ExcelPivotTableFieldNumericGroup.GroupIntervalPath);
			}
			private set
			{
				base.SetXmlNodeString(ExcelPivotTableFieldNumericGroup.GroupIntervalPath, value.ToString(CultureInfo.InvariantCulture));
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTableFieldNumericGroup"/>.
		/// </summary>
		/// <param name="ns">The namespace of the worksheet.</param>
		/// <param name="topNode">The top node in the xml.</param>
		internal ExcelPivotTableFieldNumericGroup(XmlNamespaceManager ns, XmlNode topNode) :
			 base(ns, topNode)
		{

		}
		#endregion
	}
}