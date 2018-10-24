using System;
using System.IO;

namespace EPPlusTest.TestHelpers
{
	internal class TempTestFile : IDisposable
	{
		#region Properties
		public FileInfo File { get; private set; }
		#endregion

		#region Constructors
		public TempTestFile(string fileName = null)
		{
			this.File = new FileInfo(fileName ?? Path.GetTempFileName());
			if (this.File.Exists)
				this.File.Delete();
		}
		#endregion

		#region IDisposable Implementation
		public void Dispose()
		{
			if (this.File.Exists)
				this.File.Delete();
		}
		#endregion
	}
}
