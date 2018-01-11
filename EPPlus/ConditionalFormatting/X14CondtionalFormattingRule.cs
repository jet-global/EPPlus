using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
	/// <summary>
	/// A class that provides access to a particular X14 (extension-data) conditional formatting rule's sqref address and other properties.
	/// </summary>
	public class X14CondtionalFormattingRule : XmlHelper
	{
		#region Class Variables
		private List<X14ConditionalFormattingFormula> _formulae;
		#endregion

		#region Properties
		/// <summary>
		/// Gets or sets the address that this Conditional Formatting rule applies to.
		/// </summary>
		public string Address
		{
			get
			{
				var sqref = this.GetXmlNodeString("xm:sqref");
				return sqref.Replace(" ", ",");
			}
			set
			{
				this.SetXmlNodeString("xm:sqref", new ExcelAddress(value).AddressSpaceSeparated);
			}
		}

		/// <summary>
		/// Gets a collection of all xm:f elements (formulas) present below this node.
		/// </summary>
		public List<X14ConditionalFormattingFormula> Formulae
		{
			get
			{
				if (_formulae == null)
				{
					_formulae = new List<X14ConditionalFormattingFormula>();
					var formulaNodes = this.TopNode.SelectNodes(@"x14:cfRule//xm:f", this.NameSpaceManager);
					foreach (XmlNode node in formulaNodes)
					{
						_formulae.Add(new X14ConditionalFormattingFormula(this.NameSpaceManager, node));
					}
				}
				return _formulae;
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Create a new wrapper for an existing x14:conditionalFormatting (Extension-data) conditional formatting block.
		/// </summary>
		/// <param name="node">An x14:conditionalFormatting node.</param>
		/// <param name="nsm">A namespace manager that should include the x14, xm, and default namespaces.</param>
		internal X14CondtionalFormattingRule(XmlNode node, XmlNamespaceManager nsm) : base(nsm, node) { }
		#endregion
	}
}