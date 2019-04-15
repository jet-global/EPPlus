using System;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Represents the rangePr XML element.
	/// </summary>
	public class RangePr : XmlHelper
	{
		#region Properties
		/// <summary>
		/// Gets or sets the @startDate xml attribute value.
		/// </summary>
		public DateTime? StartDate
		{
			get
			{
				string value = base.GetXmlNodeString("@startDate");
				if (string.IsNullOrEmpty(value))
					return null;
				return DateTime.Parse(value);
			}
			set { base.SetXmlNodeString("@startDate", ConvertUtil.ConvertObjectToXmlAttributeString(value), true); }
		}

		/// <summary>
		/// Gets or sets the @endDate xml attribute value.
		/// </summary>
		public DateTime? EndDate
		{
			get
			{
				string value = base.GetXmlNodeString("@endDate");
				if (string.IsNullOrEmpty(value))
					return null;
				return DateTime.Parse(value);
			}
			set { base.SetXmlNodeString("@endDate", ConvertUtil.ConvertObjectToXmlAttributeString(value), true); }
		}

		/// <summary>
		/// Gets or sets the type of group by that this grouping represents.
		/// </summary>
		public string GroupBy
		{
			get { return base.GetXmlNodeString("@groupBy"); }
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Constructor.
		/// </summary>
		/// <param name="ns">The namespace manager to use for this element.</param>
		/// <param name="topNode">The topnode that represents this xml element.</param>
		public RangePr(XmlNamespaceManager ns, XmlNode topNode) : base(ns, topNode)
		{
			if (ns == null)
				throw new ArgumentNullException(nameof(ns));
			if (topNode == null)
				throw new ArgumentNullException(nameof(topNode));
		}
		#endregion
	}
}
