using System;
using System.IO;

namespace EPPlusTest.TestHelpers
{
	/// <summary>
	/// Test helper for temporary files.
	/// </summary>
	internal class TempTestFile : IDisposable
	{
		#region Properties
		/// <summary>
		/// Gets the file.
		/// </summary>
		public FileInfo File { get; private set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Constructor for using temporary test files.
		/// </summary>
		/// <param name="tempFileName">The temp file name to use (optional--Note that this file WILL BE DELETED).</param>
		/// <remarks>The <paramref name="tempFileName"/> file WILL BE DELETED.</remarks>
		public TempTestFile(string tempFileName = null)
		{
			this.File = new FileInfo(tempFileName ?? Path.GetTempFileName());
			if (this.File.Exists)
				this.File.Delete();
		}
		#endregion

		#region IDisposable Implementation
		/// <summary>
		/// Deletes the temporary file.
		/// </summary>
		public void Dispose()
		{
			if (this.File.Exists)
				this.File.Delete();
		}
		#endregion
	}
}
