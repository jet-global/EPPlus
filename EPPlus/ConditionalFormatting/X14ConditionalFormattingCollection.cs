using System;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
	/// <summary>
	/// A class that allows access to X14 (extension-data) conditional formatting rules.
	/// </summary>
	public class X14ConditionalFormattingCollection : XmlHelper
	{
		#region Properties
		/// <summary>
		/// Gets a list of <see cref="X14CondtionalFormattingRule"/>s contained on this collection's worksheet.
		/// </summary>
		public List<X14CondtionalFormattingRule> X14Rules { get; } = new List<X14CondtionalFormattingRule>();
		#endregion

		#region Constructors
		/// <summary>
		/// Initialize a wrapper around existing x14 conditional formatting rules that provides access to sqref and formula sub-nodes.
		/// </summary>
		/// <param name="worksheet">The worksheet whose collection of x14 conditional formatting nodes should be modeled.</param>
		internal X14ConditionalFormattingCollection(ExcelWorksheet worksheet) : base(worksheet.NameSpaceManager, worksheet.WorksheetXml.DocumentElement)
		{
			string x14CFpath = @"d:worksheet/d:extLst/d:ext/x14:conditionalFormattings/x14:conditionalFormatting";
			var nodes = worksheet.WorksheetXml.SelectNodes(x14CFpath, worksheet.NameSpaceManager);
			foreach (XmlNode node in nodes)
			{
				var x14Rule = new X14CondtionalFormattingRule(node, worksheet.NameSpaceManager);
				if (!string.IsNullOrEmpty(x14Rule.Address))
					this.X14Rules.Add(x14Rule);
			}
		}
		#endregion

		#region Virtual Methods
		/// <summary>
		/// Applies the <paramref name="transformer"/> to all formulas in the <see cref="X14ConditionalFormattingCollection"/>.
		/// </summary>
		/// <param name="transformer">The transformation to apply.</param>
		public virtual void TransformFormulaReferences(Func<string, string> transformer)
		{
			XmlHelper.TransformValuesInNode(this.TopNode, this.NameSpaceManager, transformer, ".//x14:cfRule//xm:f");
		}
		#endregion
	}
}
