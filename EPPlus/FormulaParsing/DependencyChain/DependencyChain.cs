using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing
{
	/// <summary>
	/// Represents a chain of dependent cells.
	/// </summary>
	internal class DependencyChain
	{
		#region Properties
		internal List<FormulaCell> List { get; } = new List<FormulaCell>();
		internal Dictionary<ulong, int> Index { get; } = new Dictionary<ulong, int>();
		internal List<int> CalcOrder { get; } = new List<int>();
		#endregion

		#region Constructors
		internal void Add(FormulaCell f)
		{
			List.Add(f);
			f.Index = List.Count - 1;
			Index.Add(ExcelCellBase.GetCellID(f.SheetID, f.Row, f.Column), f.Index);
		}
		#endregion
	}
}