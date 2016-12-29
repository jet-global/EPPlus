using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting
{
	/// <summary>
	/// A wrapper for providing access to an xm:f subnode of an x14 conditional formatting rule.
	/// </summary>
	public class X14ConditionalFormattingFormula : XmlHelper, IExcelConditionalFormattingWithFormula
	{
		#region Properties
		/// <summary>
		/// Gets or sets the value of the formula this object represents.
		/// </summary>
		public string Formula
		{
			get
			{
				return this.TopNode.InnerText;
			}
			set
			{
				this.TopNode.InnerText = value;
			}
		}
		#endregion

		#region Constructor
		/// <summary>
		/// Create an object that represents an xm:f formula node.
		/// </summary>
		/// <param name="nameSpaceManager">A namespace manager that includes the xm namespace.</param>
		/// <param name="xmfNode">The xm:f node whose formula should be wrapped.</param>
		internal X14ConditionalFormattingFormula(XmlNamespaceManager nameSpaceManager, XmlNode xmfNode) : base(nameSpaceManager, xmfNode) { }
		#endregion
	}
}
