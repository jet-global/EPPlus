using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// A pivot table field date group.
	/// </summary>
	public class ExcelPivotTableFieldDateGroup : ExcelPivotTableFieldGroup
	{
		#region Constants
		private const string GroupByPath = "d:fieldGroup/d:rangePr/@groupBy";
		#endregion

		#region Properties
		/// <summary>
		/// Gets how to group the date field.
		/// </summary>
		public eDateGroupBy GroupBy
		{
			get
			{
				string v = base.GetXmlNodeString(ExcelPivotTableFieldDateGroup.GroupByPath);
				if (v != "")
					return (eDateGroupBy)Enum.Parse(typeof(eDateGroupBy), v, true);
				else
					throw (new Exception("Invalid date Groupby"));
			}
			private set
			{
				base.SetXmlNodeString(ExcelPivotTableFieldDateGroup.GroupByPath, value.ToString().ToLower(CultureInfo.InvariantCulture));
			}
		}

		/// <summary>
		/// Gets whether there exist an auto detect start date.
		/// </summary>
		public bool AutoStart
		{
			get
			{
				return base.GetXmlNodeBool("@autoStart", false);
			}
		}

		/// <summary>
		/// Gets whether there exist an auto detect end date.
		/// </summary>
		public bool AutoEnd
		{
			get
			{
				return base.GetXmlNodeBool("@autoStart", false);
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTableFieldDateGroup"/>.
		/// </summary>
		/// <param name="ns">The namespace of the worksheet.</param>
		/// <param name="topNode">The top node in the xml.</param>
		internal ExcelPivotTableFieldDateGroup(XmlNamespaceManager ns, XmlNode topNode) :
			 base(ns, topNode)
		{

		}
		#endregion
	}
}