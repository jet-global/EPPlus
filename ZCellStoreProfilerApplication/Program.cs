using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ZCellStoreProfilerApplication
{
	class Program
	{
		static void Main(string[] args)
		{

			var random = new Random(); // TODO seed
			var cellStoreProfiler = new ZCellStoreProfiler<int>();
			for (int i = 0; i < 200; i++)
			{
				switch (random.Next(25))
				{
					case 0:
					case 21:
					case 22:
					case 23:
					case 24:
						Console.Write("Enumerating... ");
						cellStoreProfiler.EnumerateItems();
						Console.WriteLine($"ZCellStore Completed in: {cellStoreProfiler.ZCellStoreTimer.ElapsedTicks}; CellStore Completed in: {cellStoreProfiler.CellStoreTimer.ElapsedTicks}");
						break;
					case 1:
						Console.Write("Clearing... ");
						cellStoreProfiler.Clear(random.Next(ExcelPackage.MaxRows), random.Next(ExcelPackage.MaxColumns), random.Next(ExcelPackage.MaxRows), random.Next(ExcelPackage.MaxColumns));
						Console.WriteLine($"ZCellStore Completed in: {cellStoreProfiler.ZCellStoreTimer.ElapsedTicks}; CellStore Completed in: {cellStoreProfiler.CellStoreTimer.ElapsedTicks}");
						break;
					case 2:
						Console.Write("Deleting... ");
						cellStoreProfiler.Delete(random.Next(ExcelPackage.MaxRows), random.Next(ExcelPackage.MaxColumns), random.Next(ExcelPackage.MaxRows), random.Next(ExcelPackage.MaxColumns));
						Console.WriteLine($"ZCellStore Completed in: {cellStoreProfiler.ZCellStoreTimer.ElapsedTicks}; CellStore Completed in: {cellStoreProfiler.CellStoreTimer.ElapsedTicks}");
						break;
					//case 3:
					//	Console.Write("Deleting with shift... ");
					//	cellStoreProfiler.Delete(random.Next(ExcelPackage.MaxRows), random.Next(ExcelPackage.MaxColumns), random.Next(ExcelPackage.MaxRows), random.Next(ExcelPackage.MaxColumns), random.Next(2) == 1);
					//	Console.WriteLine($"ZCellStore Completed in: {cellStoreProfiler.ZCellStoreTimer.ElapsedTicks}; CellStore Completed in: {cellStoreProfiler.CellStoreTimer.ElapsedTicks}");
					//	break;
					case 4:
						Console.Write("Exists with value... ");
						cellStoreProfiler.Exists(random.Next(ExcelPackage.MaxRows), random.Next(ExcelPackage.MaxColumns), out _);
						Console.WriteLine($"ZCellStore Completed in: {cellStoreProfiler.ZCellStoreTimer.ElapsedTicks}; CellStore Completed in: {cellStoreProfiler.CellStoreTimer.ElapsedTicks}");
						break;
					case 5:
						Console.Write("Exists... ");
						cellStoreProfiler.Exists(random.Next(ExcelPackage.MaxRows), random.Next(ExcelPackage.MaxColumns));
						Console.WriteLine($"ZCellStore Completed in: {cellStoreProfiler.ZCellStoreTimer.ElapsedTicks}; CellStore Completed in: {cellStoreProfiler.CellStoreTimer.ElapsedTicks}");
						break;
					case 6:
						Console.Write("GetDimension... ");
						cellStoreProfiler.GetDimension(out _, out _, out _, out _);
						Console.WriteLine($"ZCellStore Completed in: {cellStoreProfiler.ZCellStoreTimer.ElapsedTicks}; CellStore Completed in: {cellStoreProfiler.CellStoreTimer.ElapsedTicks}");
						break;
					case 7:
						Console.Write("GetValue... ");
						cellStoreProfiler.GetValue(random.Next(ExcelPackage.MaxRows), random.Next(ExcelPackage.MaxColumns));
						Console.WriteLine($"ZCellStore Completed in: {cellStoreProfiler.ZCellStoreTimer.ElapsedTicks}; CellStore Completed in: {cellStoreProfiler.CellStoreTimer.ElapsedTicks}");
						break;
					case 8:
						Console.Write("Insert... ");
						cellStoreProfiler.Insert(random.Next(ExcelPackage.MaxRows), random.Next(ExcelPackage.MaxColumns), random.Next(ExcelPackage.MaxRows), random.Next(ExcelPackage.MaxColumns));
						Console.WriteLine($"ZCellStore Completed in: {cellStoreProfiler.ZCellStoreTimer.ElapsedTicks}; CellStore Completed in: {cellStoreProfiler.CellStoreTimer.ElapsedTicks}");
						break;
					case 9:
						Console.Write("NextCell... ");
						var startRow = random.Next(ExcelPackage.MaxRows);
						var startColumn = random.Next(ExcelPackage.MaxColumns);
						cellStoreProfiler.NextCell(ref startRow, ref startColumn);
						Console.WriteLine($"ZCellStore Completed in: {cellStoreProfiler.ZCellStoreTimer.ElapsedTicks}; CellStore Completed in: {cellStoreProfiler.CellStoreTimer.ElapsedTicks}");
						break;
					case 10:
						Console.Write("PrevCell... ");
						var startRow2 = random.Next(ExcelPackage.MaxRows);
						var startColumn2 = random.Next(ExcelPackage.MaxColumns);
						cellStoreProfiler.PrevCell(ref startRow2, ref startColumn2);
						Console.WriteLine($"ZCellStore Completed in: {cellStoreProfiler.ZCellStoreTimer.ElapsedTicks}; CellStore Completed in: {cellStoreProfiler.CellStoreTimer.ElapsedTicks}");
						break;
					case 11:
					case 12:
					case 13:
					case 14:
					case 15:
					case 16:
					case 17:
					case 18:
					case 19:
					case 20:
						Console.Write("SetValue... ");
						cellStoreProfiler.SetValue(random.Next(ExcelPackage.MaxRows), random.Next(ExcelPackage.MaxColumns), random.Next());
						Console.WriteLine($"ZCellStore Completed in: {cellStoreProfiler.ZCellStoreTimer.ElapsedTicks}; CellStore Completed in: {cellStoreProfiler.CellStoreTimer.ElapsedTicks}");
						break;
				}
			}
			Console.WriteLine("Saving results... ");
			cellStoreProfiler.SaveResults(@"C:\profiler.csv");
		}
	}
}
