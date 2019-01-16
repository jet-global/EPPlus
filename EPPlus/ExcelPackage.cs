/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied.
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 *
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		                Initial Release		        2009-10-01
 * Starnuto Di Topo & Jan Källman   Added stream constructors
 *                                  and Load method Save as
 *                                  stream                      2010-03-14
 * Jan Källman		License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.IO;
using System.Resources;
using System.Security.Cryptography;
using System.Xml;
using OfficeOpenXml.Encryption;
using OfficeOpenXml.Internationalization;
using OfficeOpenXml.Utils;
namespace OfficeOpenXml
{
	#region Enumerators
	/// <summary>
	/// Maps to DotNetZips CompressionLevel enum
	/// </summary>
	public enum CompressionLevel
	{
		Level0 = 0,
		None = 0,
		Level1 = 1,
		BestSpeed = 1,
		Level2 = 2,
		Level3 = 3,
		Level4 = 4,
		Level5 = 5,
		Level6 = 6,
		Default = 6,
		Level7 = 7,
		Level8 = 8,
		BestCompression = 9,
		Level9 = 9,
	}
	#endregion

	/// <summary>
	/// Represents an Excel 2007+ XLSX file package.
	/// This is the top-level object to access all parts of the document.
	/// Code samples can be found at  <a href="http://epplus.codeplex.com/">http://epplus.codeplex.com/</a>
	/// </summary>
	public sealed class ExcelPackage : IDisposable
	{
		#region Constants
		/// <summary>
		/// Extention Schema types
		/// </summary>
		internal const string schemaXmlExtension = "application/xml";
		internal const string schemaRelsExtension = "application/vnd.openxmlformats-package.relationships+xml";
		/// <summary>
		/// Main Xml schema name
		/// </summary>
		internal const string schemaMain = @"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
		internal const string schemaMain2009 = @"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
		internal const string schemaOfficeMain2006 = @"http://schemas.microsoft.com/office/excel/2006/main";
		/// <summary>
		/// Relationship schema name
		/// </summary>
		internal const string schemaRelationships = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships";

		internal const string schemaDrawings = @"http://schemas.openxmlformats.org/drawingml/2006/main";
		internal const string schemaSheetDrawings = @"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
		internal const string schemaMarkupCompatibility = @"http://schemas.openxmlformats.org/markup-compatibility/2006";

		internal const string schemaMicrosoftVml = @"urn:schemas-microsoft-com:vml";
		internal const string schemaMicrosoftOffice = "urn:schemas-microsoft-com:office:office";
		internal const string schemaMicrosoftExcel = "urn:schemas-microsoft-com:office:excel";

		internal const string schemaChart = @"http://schemas.openxmlformats.org/drawingml/2006/chart";
		internal const string schemaHyperlink = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
		internal const string schemaComment = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
		internal const string schemaImage = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
		internal const string schemaSlicerDrawing = @"http://schemas.microsoft.com/office/drawing/2010/slicer";
		internal const string schemaSlicerRelationship = @"http://schemas.microsoft.com/office/2007/relationships/slicer";
		internal const string schemaSlicerCache = @"http://schemas.microsoft.com/office/2007/relationships/slicerCache";

		//Office properties
		internal const string schemaCore = @"http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
		internal const string schemaExtended = @"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
		internal const string schemaCustom = @"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
		internal const string schemaDc = @"http://purl.org/dc/elements/1.1/";
		internal const string schemaDcTerms = @"http://purl.org/dc/terms/";
		internal const string schemaDcmiType = @"http://purl.org/dc/dcmitype/";
		internal const string schemaXsi = @"http://www.w3.org/2001/XMLSchema-instance";
		internal const string schemaVt = @"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

		//Pivottables
		internal const string schemaPivotCacheRelationship = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
		internal const string schemaPivotCacheRecordsRelationship = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
		internal const string schemaPivotTable = @"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";
		internal const string schemaPivotCacheDefinition = @"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
		internal const string schemaPivotCacheRecords = @"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";

		//VBA
		internal const string schemaVBA = @"application/vnd.ms-office.vbaProject";
		internal const string schemaVBASignature = @"application/vnd.ms-office.vbaProjectSignature";

		internal const string contentTypeWorkbookDefault = @"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
		internal const string contentTypeWorkbookMacroEnabled = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
		internal const string contentTypeSharedString = @"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";

		/// <summary>
		/// Maximum number of columns in a worksheet (16384).
		/// </summary>
		public const int MaxColumns = 16384;
		/// <summary>
		/// Maximum number of rows in a worksheet (1048576).
		/// </summary>
		public const int MaxRows = 1048576;

		internal const bool preserveWhitespace = false;
		#endregion

		#region Class Variables
		private Stream _stream = null;
		private bool _isExternalStream = false;
		private Packaging.ZipPackage _package;
		private ExcelWorkbook _workbook;
		private IFormulaManager _FormulaManager;
		private ExcelEncryption _encryption = null;
		private FileInfo _file = null;
		private static object _lock = new object();
		#endregion

		#region Nested Classes
		internal class ImageInfo
		{
			internal string Hash { get; set; }
			internal Uri Uri { get; set; }
			internal int RefCount { get; set; }
			internal Packaging.ZipPackagePart Part { get; set; }
		}
		#endregion

		#region Properties
		/// <summary>
		/// Gets or sets the collection of Images contained in the package.
		/// </summary>
		internal Dictionary<string, ImageInfo> Images { get; set; } = new Dictionary<string, ImageInfo>();

		/// <summary>
		/// Gets the <see cref="IFormulaManager"/> for this <see cref="ExcelPackage"/> that
		/// can be used to update formulas in cells and charts.
		/// </summary>
		public IFormulaManager FormulaManager
		{
			get
			{
				if (this._FormulaManager == null)
					this._FormulaManager = new FormulaManager();
				return this._FormulaManager;
			}
			private set
			{
				this._FormulaManager = value;
			}
		}

		/// <summary>
		/// Gets the <see cref="Packaging.ZipPackage"/> that this <see cref="ExcelPackage"/> represents.
		/// </summary>
		public Packaging.ZipPackage Package { get { return (this._package); } }

		/// <summary>
		/// Information how and if the package is encrypted
		/// </summary>
		public ExcelEncryption Encryption
		{
			get
			{
				if (this._encryption == null)
				{
					this._encryption = new ExcelEncryption();
				}
				return this._encryption;
			}
		}

		/// <summary>
		/// Returns a reference to the workbook component within the package.
		/// All worksheets and cells can be accessed through the workbook.
		/// </summary>
		public ExcelWorkbook Workbook
		{
			get
			{
				if (this._workbook == null)
				{
					var nsm = this.CreateDefaultNSM();

					this._workbook = new ExcelWorkbook(this, nsm);
					this._workbook.GetDefinedNames();

				}
				return (this._workbook);
			}
		}

		/// <summary>
		/// The output file. Null if no file is used
		/// </summary>
		public FileInfo File
		{
			get
			{
				return this._file;
			}
			set
			{
				this._file = value;
			}
		}

		/// <summary>
		/// The output stream. This stream is the not the encrypted package.
		/// To get the encrypted package use the SaveAs(stream) method.
		/// </summary>
		public Stream Stream
		{
			get
			{
				return this._stream;
			}
		}

		/// <summary>
		/// Automaticlly adjust drawing size when column width/row height are adjusted, depending on the drawings editBy property.
		/// Default True
		/// </summary>
		public bool DoAdjustDrawings
		{
			get;
			set;
		}

		/// <summary>
		/// Compression option for the package
		/// </summary>
		public CompressionLevel Compression
		{
			get
			{
				return this.Package.Compression;
			}
			set
			{
				this.Package.Compression = value;
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Create a new instance of the ExcelPackage. Output is accessed through the Stream property.
		/// </summary>
		public ExcelPackage()
		{
			this.Init();
			this.ConstructNewFile(null);
		}

		/// <summary>
		/// Create a new instance of the ExcelPackage class based on a existing file or creates a new file.
		/// </summary>
		/// <param name="newFile">If newFile exists, it is opened.  Otherwise it is created from scratch.</param>
		public ExcelPackage(FileInfo newFile)
		{
			this.Init();
			this.File = newFile;
			this.ConstructNewFile(null);
		}
		/// <summary>
		/// Create a new instance of the ExcelPackage class based on a existing file or creates a new file.
		/// </summary>
		/// <param name="newFile">If newFile exists, it is opened.  Otherwise it is created from scratch.</param>
		/// <param name="password">Password for an encrypted package</param>
		public ExcelPackage(FileInfo newFile, string password)
		{
			this.Init();
			this.File = newFile;
			this.ConstructNewFile(password);
		}
		/// <summary>
		/// Create a new instance of the ExcelPackage class based on a existing template.
		/// If newFile exists, it will be overwritten when the Save method is called
		/// </summary>
		/// <param name="newFile">The name of the Excel file to be created</param>
		/// <param name="template">The name of the Excel template to use as the basis of the new Excel file</param>
		public ExcelPackage(FileInfo newFile, FileInfo template)
		{
			this.Init();
			this.File = newFile;
			this.CreateFromTemplate(template, null);
		}
		/// <summary>
		/// Create a new instance of the ExcelPackage class based on a existing template.
		/// If newFile exists, it will be overwritten when the Save method is called
		/// </summary>
		/// <param name="newFile">The name of the Excel file to be created</param>
		/// <param name="template">The name of the Excel template to use as the basis of the new Excel file</param>
		/// <param name="password">Password to decrypted the template</param>
		public ExcelPackage(FileInfo newFile, FileInfo template, string password)
		{
			this.Init();
			this.File = newFile;
			this.CreateFromTemplate(template, password);
		}
		/// <summary>
		/// Create a new instance of the ExcelPackage class based on a existing template.
		/// </summary>
		/// <param name="template">The name of the Excel template to use as the basis of the new Excel file</param>
		/// <param name="useStream">if true use a stream. If false create a file in the temp dir with a random name</param>
		public ExcelPackage(FileInfo template, bool useStream)
		{
			this.Init();
			this.CreateFromTemplate(template, null);
			if (useStream == false)
			{
				this.File = new FileInfo(Path.GetTempPath() + Guid.NewGuid().ToString() + ".xlsx");
			}
		}
		/// <summary>
		/// Create a new instance of the ExcelPackage class based on a existing template.
		/// </summary>
		/// <param name="template">The name of the Excel template to use as the basis of the new Excel file</param>
		/// <param name="useStream">if true use a stream. If false create a file in the temp dir with a random name</param>
		/// <param name="password">Password to decrypted the template</param>
		public ExcelPackage(FileInfo template, bool useStream, string password)
		{
			this.Init();
			this.CreateFromTemplate(template, password);
			if (useStream == false)
			{
				this.File = new FileInfo(Path.GetTempPath() + Guid.NewGuid().ToString() + ".xlsx");
			}
		}
		/// <summary>
		/// Create a new instance of the ExcelPackage class based on a stream
		/// </summary>
		/// <param name="newStream">The stream object can be empty or contain a package. The stream must be Read/Write</param>
		public ExcelPackage(Stream newStream)
		{
			this.Init();
			if (newStream.Length == 0)
			{
				this._stream = newStream;
				this._isExternalStream = true;
				this.ConstructNewFile(null);
			}
			else
			{
				this.Load(newStream);
			}
		}
		/// <summary>
		/// Create a new instance of the ExcelPackage class based on a stream
		/// </summary>
		/// <param name="newStream">The stream object can be empty or contain a package. The stream must be Read/Write</param>
		/// <param name="Password">The password to decrypt the document</param>
		public ExcelPackage(Stream newStream, string Password)
		{
			if (!(newStream.CanRead && newStream.CanWrite))
			{
				throw new Exception("The stream must be read/write");
			}

			this.Init();
			if (newStream.Length > 0)
			{
				this.Load(newStream, Password);
			}
			else
			{
				this._stream = newStream;
				this._isExternalStream = true;
				this._package = new Packaging.ZipPackage(_stream);
				this.CreateBlankWb();
			}
		}
		/// <summary>
		/// Create a new instance of the ExcelPackage class based on a stream
		/// </summary>
		/// <param name="newStream">The output stream. Must be an empty read/write stream.</param>
		/// <param name="templateStream">This stream is copied to the output stream at load</param>
		public ExcelPackage(Stream newStream, Stream templateStream)
		{
			if (newStream.Length > 0)
			{
				throw (new Exception("The output stream must be empty. Length > 0"));
			}
			else if (!(newStream.CanRead && newStream.CanWrite))
			{
				throw new Exception("The stream must be read/write");
			}
			this.Init();
			this.Load(templateStream, newStream, null);
		}
		/// <summary>
		/// Create a new instance of the ExcelPackage class based on a stream
		/// </summary>
		/// <param name="newStream">The output stream. Must be an empty read/write stream.</param>
		/// <param name="templateStream">This stream is copied to the output stream at load</param>
		/// <param name="Password">Password to decrypted the template</param>
		public ExcelPackage(Stream newStream, Stream templateStream, string Password)
		{
			if (newStream.Length > 0)
			{
				throw (new Exception("The output stream must be empty. Length > 0"));
			}
			else if (!(newStream.CanRead && newStream.CanWrite))
			{
				throw new Exception("The stream must be read/write");
			}
			this.Init();
			this.Load(templateStream, newStream, Password);
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Closes the package.
		/// </summary>
		public void Dispose()
		{
			if (this._package != null)
			{
				if (this._isExternalStream == false && Stream != null && (this.Stream.CanRead || this.Stream.CanWrite))
				{
					this.Stream.Close();
				}
				this._package.Close();
				if (this._isExternalStream == false) ((IDisposable)this._stream).Dispose();
				if (this._workbook != null)
				{
					this._workbook.Dispose();
				}
				this._package = null;
				this.Images = null;
				this._file = null;
				this._workbook = null;
				this._stream = null;
				this._workbook = null;
				GC.Collect();
			}
		}

		/// <summary>
		/// Saves all the components back into the package.
		/// This method recursively calls the Save method on all sub-components.
		/// We close the package after the save is done.
		/// </summary>
		public void Save()
		{
			try
			{
				if (this.Stream is MemoryStream && this.Stream.Length > 0)
				{
					//Close any open memorystream and "renew" then. This can occure if the package is saved twice.
					//The stream is left open on save to enable the user to read the stream-property.
					//Non-memorystream streams will leave the closing to the user before saving a second time.
					this.CloseStream();
				}

				this.Workbook.Save();
				if (this.File == null)
				{
					if (this.Encryption.IsEncrypted)
					{
#if !MONO
						var ms = new MemoryStream();
						this.Package.Save(ms);
						byte[] file = ms.ToArray();
						EncryptedPackageHandler eph = new EncryptedPackageHandler();
						var msEnc = eph.EncryptPackage(file, Encryption);
						ExcelPackage.CopyStream(msEnc, ref _stream);
#endif
#if MONO
                        throw new NotSupportedException("Encryption is not supported under Mono.");
#endif
					}
					else
					{
						this.Package.Save(this.Stream);
					}
					this.Stream.Flush();
					this.Package.Close();
				}
				else
				{
					if (System.IO.File.Exists(File.FullName))
					{
						try
						{
							System.IO.File.Delete(File.FullName);
						}
						catch (Exception ex)
						{
							throw (new Exception(string.Format("Error overwriting file {0}", File.FullName), ex));
						}
					}

					this.Package.Save(this.Stream);
					this.Package.Close();
					if (this.Stream is MemoryStream)
					{
						var fi = new FileStream(this.File.FullName, FileMode.Create);
						//EncryptPackage
						if (this.Encryption.IsEncrypted)
						{
#if !MONO
							byte[] file = ((MemoryStream)this.Stream).ToArray();
							EncryptedPackageHandler eph = new EncryptedPackageHandler();
							var ms = eph.EncryptPackage(file, Encryption);

							fi.Write(ms.GetBuffer(), 0, (int)ms.Length);
#endif
#if MONO
                            throw new NotSupportedException("Encryption is not supported under Mono.");
#endif
						}
						else
						{
							fi.Write(((MemoryStream)this.Stream).GetBuffer(), 0, (int)this.Stream.Length);
						}
						fi.Close();
					}
					else
					{
						System.IO.File.WriteAllBytes(this.File.FullName, this.GetAsByteArray(false));
					}
				}
			}
			catch (Exception ex)
			{
				if (this.File == null)
				{
					throw;
				}
				else
				{
					throw (new InvalidOperationException(string.Format("Error saving file {0}", this.File.FullName), ex));
				}
			}
		}

		/// <summary>
		/// Saves all the components back into the package.
		/// This method recursively calls the Save method on all sub-components.
		/// The package is closed after it has ben saved
		/// d to encrypt the workbook with.
		/// </summary>
		/// <param name="password">This parameter overrides the Workbook.Encryption.Password.</param>
		public void Save(string password)
		{
			this.Encryption.Password = password;
			this.Save();
		}

		/// <summary>
		/// Saves the workbook to a new file. The package is closed after it has been saved.
		/// </summary>
		/// <param name="fileName">The filename to save the workbook to.</param>
		public void SaveAs(string fileName)
		{
			this.SaveAs(new FileInfo(fileName));
		}

		/// <summary>
		/// Saves the workbook to a new file
		/// The package is closed after it has been saved
		/// </summary>
		/// <param name="file">The file location</param>
		public void SaveAs(FileInfo file)
		{
			this.File = file;
			this.Save();
		}

		/// <summary>
		/// Saves the workbook to a new file
		/// The package is closed after it has been saved
		/// </summary>
		/// <param name="file">The file</param>
		/// <param name="password">The password to encrypt the workbook with.
		/// This parameter overrides the Encryption.Password.</param>
		public void SaveAs(FileInfo file, string password)
		{
			this.File = file;
			this.Encryption.Password = password;
			this.Save();
		}

		/// <summary>
		/// Copies the Package to the Outstream
		/// The package is closed after it has been saved
		/// </summary>
		/// <param name="outputStream">The stream to copy the package to</param>
		public void SaveAs(Stream outputStream)
		{
			this.File = null;
			this.Save();

			if (outputStream != _stream)
			{
				if (this.Encryption.IsEncrypted)
				{
#if !MONO
					//Encrypt Workbook
					Byte[] file = new byte[this.Stream.Length];
					long pos = this.Stream.Position;
					this.Stream.Seek(0, SeekOrigin.Begin);
					this.Stream.Read(file, 0, (int)this.Stream.Length);
					EncryptedPackageHandler eph = new EncryptedPackageHandler();
					var ms = eph.EncryptPackage(file, this.Encryption);
					ExcelPackage.CopyStream(ms, ref outputStream);
#endif
#if MONO
                throw new NotSupportedException("Encryption is not supported under Mono.");
#endif
				}
				else
				{
					ExcelPackage.CopyStream(_stream, ref outputStream);
				}
			}
		}

		/// <summary>
		/// Copies the Package to the Outstream
		/// The package is closed after it has been saved
		/// </summary>
		/// <param name="outputStream">The stream to copy the package to</param>
		/// <param name="password">The password to encrypt the workbook with.
		/// This parameter overrides the Encryption.Password.</param>
		public void SaveAs(Stream outputStream, string password)
		{
			this.Encryption.Password = password;
			this.SaveAs(outputStream);
		}

		/// <summary>
		/// Saves and returns the Excel files as a bytearray.
		/// Note that the package is closed upon save
		/// </summary>
		/// <returns>The .xlsx file as a byte array.</returns>
		public byte[] GetAsByteArray()
		{
			return this.GetAsByteArray(true);
		}

		/// <summary>
		/// Saves and returns the Excel files as a bytearray
		/// Note that the package is closed upon save
		/// </summary>
		/// <param name="password">The password to encrypt the workbook with.
		/// This parameter overrides the Encryption.Password.</param>
		/// <returns>The encrypted xlsx file, as a byte array.</returns>
		public byte[] GetAsByteArray(string password)
		{
			if (password != null)
			{
				this.Encryption.Password = password;
			}
			return this.GetAsByteArray(true);
		}

		/// <summary>
		/// Loads the specified package data from a stream.
		/// </summary>
		/// <param name="input">The input.</param>
		public void Load(Stream input)
		{
			this.Load(input, new MemoryStream(), null);
		}

		/// <summary>
		/// Loads the specified package data from a stream.
		/// </summary>
		/// <param name="input">The input.</param>
		/// <param name="Password">The password to decrypt the document</param>
		public void Load(Stream input, string Password)
		{
			this.Load(input, new MemoryStream(), Password);
		}

		/// <summary>
		/// Configures this <see cref="ExcelPackage"/> instance with the given <paramref name="formulaManager"/>.
		/// </summary>
		/// <param name="formulaManager">The <see cref="IFormulaManager"/> to use when updating formulas.</param>
		public void Configure(IFormulaManager formulaManager)
		{
			this._FormulaManager = formulaManager;
		}
		#endregion

		#region Internal Static Methods
		/// <summary>
		/// Copies the input stream to the output stream.
		/// </summary>
		/// <param name="inputStream">The input stream.</param>
		/// <param name="outputStream">The output stream.</param>
		internal static void CopyStream(Stream inputStream, ref Stream outputStream)
		{
			if (!inputStream.CanRead)
			{
				throw (new Exception("Can not read from inputstream"));
			}
			if (!outputStream.CanWrite)
			{
				throw (new Exception("Can not write to outputstream"));
			}
			if (inputStream.CanSeek)
			{
				inputStream.Seek(0, SeekOrigin.Begin);
			}

			const int bufferLength = 8096;
			var buffer = new Byte[bufferLength];
			lock (_lock)
			{
				int bytesRead = inputStream.Read(buffer, 0, bufferLength);
				// write the required bytes
				while (bytesRead > 0)
				{
					outputStream.Write(buffer, 0, bytesRead);
					bytesRead = inputStream.Read(buffer, 0, bufferLength);
				}
				outputStream.Flush();
			}
		}
		#endregion

		#region Internal Methods
		/// <summary>
		/// Get the contents of this package as a byte array.
		/// </summary>
		/// <param name="save">True if the workbook should be saved first before retrieving the bytes.</param>
		/// <returns>The package contents as a byte array.</returns>
		internal byte[] GetAsByteArray(bool save)
		{
			if (save)
			{
				this.Workbook.Save();
				this.Package.Close();
				this.Package.Save(this.Stream);
			}
			Byte[] byRet = new byte[Stream.Length];
			long pos = this.Stream.Position;
			this.Stream.Seek(0, SeekOrigin.Begin);
			this.Stream.Read(byRet, 0, (int)this.Stream.Length);

			//Encrypt Workbook?
			if (this.Encryption.IsEncrypted)
			{
#if !MONO
				EncryptedPackageHandler eph = new EncryptedPackageHandler();
				var ms = eph.EncryptPackage(byRet, this.Encryption);
				byRet = ms.ToArray();
#endif
			}

			this.Stream.Seek(pos, SeekOrigin.Begin);
			this.Stream.Close();
			return byRet;
		}

		/// <summary>
		/// Add an image to the workbook.
		/// </summary>
		/// <param name="image">The byte array representation of the image to add.</param>
		/// <returns>Information about the newly-added image.</returns>
		internal ImageInfo AddImage(byte[] image)
		{
			return this.AddImage(image, null, "");
		}

		/// <summary>
		/// Add an image to the workbook.
		/// </summary>
		/// <param name="image">The image to add in byte-array form.</param>
		/// <param name="uri">The Uri to store the image at.</param>
		/// <param name="contentType">The content type of the image.</param>
		/// <returns>Information about the newly-added image.</returns>
		internal ImageInfo AddImage(byte[] image, Uri uri, string contentType)
		{
			var hashProvider = new SHA1CryptoServiceProvider();
			var hash = BitConverter.ToString(hashProvider.ComputeHash(image)).Replace("-", "");
			lock (this.Images)
			{
				if (this.Images.ContainsKey(hash))
				{
					this.Images[hash].RefCount++;
				}
				else
				{
					Packaging.ZipPackagePart imagePart;
					if (uri == null)
					{
						uri = GetNewUri(Package, "/xl/media/image{0}.jpg");
						imagePart = Package.CreatePart(uri, "image/jpeg", CompressionLevel.None);
					}
					else
					{
						imagePart = Package.CreatePart(uri, contentType, CompressionLevel.None);
					}
					var stream = imagePart.GetStream(FileMode.Create, FileAccess.Write);
					stream.Write(image, 0, image.GetLength(0));

					this.Images.Add(hash, new ImageInfo() { Uri = uri, RefCount = 1, Hash = hash, Part = imagePart });
				}
			}
			return this.Images[hash];
		}

		/// <summary>
		/// Load bytes into an existing image.
		/// </summary>
		/// <param name="image">The new image bytes.</param>
		/// <param name="uri">The Uri to load to.</param>
		/// <param name="imagePart">The part to load to.</param>
		/// <returns>Information about the newly-reloaded image.</returns>
		internal ImageInfo LoadImage(byte[] image, Uri uri, Packaging.ZipPackagePart imagePart)
		{
			var hashProvider = new SHA1CryptoServiceProvider();
			var hash = BitConverter.ToString(hashProvider.ComputeHash(image)).Replace("-", "");
			if (this.Images.ContainsKey(hash))
			{
				this.Images[hash].RefCount++;
			}
			else
			{
				this.Images.Add(hash, new ImageInfo() { Uri = uri, RefCount = 1, Hash = hash, Part = imagePart });
			}
			return this.Images[hash];
		}

		/// <summary>
		/// Remove an image from the report.
		/// </summary>
		/// <param name="hash">The hash value of the image to remove.</param>
		internal void RemoveImage(string hash)
		{
			lock (this.Images)
			{
				if (this.Images.ContainsKey(hash))
				{
					var ii = Images[hash];
					ii.RefCount--;
					if (ii.RefCount == 0)
					{
						this.Package.DeletePart(ii.Uri);
						this.Images.Remove(hash);
					}
				}
			}
		}

		/// <summary>
		/// Get information about a specific image in the worksheet.
		/// </summary>
		/// <param name="image">The bytes of the image to look up.</param>
		/// <returns>Information about the image, or null if the image could not be found.</returns>
		internal ImageInfo GetImageInfo(byte[] image)
		{
			var hashProvider = new SHA1CryptoServiceProvider();
			var hash = BitConverter.ToString(hashProvider.ComputeHash(image)).Replace("-", "");

			if (this.Images.ContainsKey(hash))
			{
				return this.Images[hash];
			}
			else
			{
				return null;
			}
		}

		/// <summary>
		/// Saves the XmlDocument into the package at the specified Uri.
		/// </summary>
		/// <param name="uri">The Uri of the component</param>
		/// <param name="xmlDoc">The XmlDocument to save</param>
		internal void SavePart(Uri uri, XmlDocument xmlDoc)
		{
			Packaging.ZipPackagePart part = this.Package.GetPart(uri);
			xmlDoc.Save(part.GetStream(FileMode.Create, FileAccess.Write));
		}

		/// <summary>
		/// Saves the XmlDocument into the package at the specified Uri.
		/// </summary>
		/// <param name="uri">The Uri of the component</param>
		/// <param name="xmlDoc">The XmlDocument to save</param>
		internal void SaveWorkbook(Uri uri, XmlDocument xmlDoc)
		{
			Packaging.ZipPackagePart part = this.Package.GetPart(uri);
			if (Workbook.VbaProject == null)
			{
				if (part.ContentType != contentTypeWorkbookDefault)
				{
					part = this.Package.CreatePart(uri, contentTypeWorkbookDefault, Compression);
				}
			}
			else
			{
				if (part.ContentType != contentTypeWorkbookMacroEnabled)
				{
					var rels = part.GetRelationships();
					this.Package.DeletePart(uri);
					part = this.Package.CreatePart(uri, contentTypeWorkbookMacroEnabled);
					foreach (var rel in rels)
					{
						this.Package.DeleteRelationship(rel.Id);
						part.CreateRelationship(rel.TargetUri, rel.TargetMode, rel.RelationshipType);
					}
				}
			}
			xmlDoc.Save(part.GetStream(FileMode.Create, FileAccess.Write));
		}

		/// <summary>
		/// Get the XmlDocument from an URI
		/// </summary>
		/// <param name="uri">The Uri to the part</param>
		/// <returns>The XmlDocument</returns>
		internal XmlDocument GetXmlFromUri(Uri uri)
		{
			XmlDocument xml = new XmlDocument();
			Packaging.ZipPackagePart part = this.Package.GetPart(uri);
			XmlHelper.LoadXmlSafe(xml, part.GetStream());
			return (xml);
		}
		#endregion

		#region Private Methods
		private Uri GetNewUri(Packaging.ZipPackage package, string sUri)
		{
			int id = 1;
			Uri uri;
			do
			{
				uri = new Uri(string.Format(sUri, id++), UriKind.Relative);
			}
			while (package.PartExists(uri));
			return uri;
		}

		/// <summary>
		/// Init values here
		/// </summary>
		private void Init()
		{
			this.DoAdjustDrawings = true;
		}

		/// <summary>
		/// Create a new file from a template
		/// </summary>
		/// <param name="template">An existing xlsx file to use as a template</param>
		/// <param name="password">The password to decrypt the package.</param>
		/// <returns></returns>
		private void CreateFromTemplate(FileInfo template, string password)
		{
			if (template != null)
				template.Refresh();
			if (template.Exists)
			{
				if (_stream == null)
					_stream = new MemoryStream();
				var ms = new MemoryStream();
				if (password != null)
				{
#if !MONO
					this.Encryption.IsEncrypted = true;
					this.Encryption.Password = password;
					var encrHandler = new EncryptedPackageHandler();
					ms = encrHandler.DecryptPackage(template, this.Encryption);
					encrHandler = null;
#endif
#if MONO
	                throw (new NotImplementedException("No support for Encrypted packages in Mono"));
#endif
				}
				else
				{
					byte[] b = System.IO.File.ReadAllBytes(template.FullName);
					ms.Write(b, 0, b.Length);
				}
				try
				{
					this._package = new Packaging.ZipPackage(ms);
				}
				catch (Exception ex)
				{
#if !MONO
					if (password == null && CompoundDocument.IsStorageFile(template.FullName) == 0)
					{
						throw new Exception("Can not open the package. Package is an OLE compound document. If this is an encrypted package, please supply the password", ex);
					}
					else
					{
						throw;
					}
#endif
#if MONO
                    throw;
#endif
				}
			}
			else
				throw new Exception("Passed invalid TemplatePath to Excel Template");
			//return newFile;
		}
		private void ConstructNewFile(string password)
		{
			var ms = new MemoryStream();
			if (this._stream == null) this._stream = new MemoryStream();
			if (this.File != null) this.File.Refresh();
			if (this.File != null && this.File.Exists)
			{
				if (password != null)
				{
#if !MONO
					var encrHandler = new EncryptedPackageHandler();
					this.Encryption.IsEncrypted = true;
					this.Encryption.Password = password;
					ms = encrHandler.DecryptPackage(File, Encryption);
					encrHandler = null;
#endif
#if MONO
                    throw new NotImplementedException("No support for Encrypted packages in Mono");
#endif
				}
				else
				{
					byte[] b = System.IO.File.ReadAllBytes(this.File.FullName);
					ms.Write(b, 0, b.Length);
				}
				try
				{
					this._package = new Packaging.ZipPackage(ms);
				}
				catch (Exception ex)
				{
#if !MONO
					if (password == null && CompoundDocument.IsStorageFile(File.FullName) == 0)
					{
						throw new Exception("Can not open the package. Package is an OLE compound document. If this is an encrypted package, please supply the password", ex);
					}
					else
					{
						throw;
					}
#endif
#if MONO
                    throw;
#endif
				}
			}
			else
			{
				//_package = Package.Open(_stream, FileMode.Create, FileAccess.ReadWrite);
				this._package = new Packaging.ZipPackage(ms);
				this.CreateBlankWb();
			}
		}

		private void CreateBlankWb()
		{
			XmlDocument workbook = Workbook.WorkbookXml; // this will create the workbook xml in the package
																		// create the relationship to the main part
			this.Package.CreateRelationship(UriHelper.GetRelativeUri(new Uri("/xl", UriKind.Relative), Workbook.WorkbookUri), Packaging.TargetMode.Internal, schemaRelationships + "/officeDocument");
		}

		private XmlNamespaceManager CreateDefaultNSM()
		{
			//  Create a NamespaceManager to handle the default namespace,
			//  and create a prefix for the default namespace:
			NameTable nt = new NameTable();
			var ns = new XmlNamespaceManager(nt);
			ns.AddNamespace(string.Empty, ExcelPackage.schemaMain);
			ns.AddNamespace("d", ExcelPackage.schemaMain);
			ns.AddNamespace("r", ExcelPackage.schemaRelationships);
			ns.AddNamespace("c", ExcelPackage.schemaChart);
			ns.AddNamespace("vt", ExcelPackage.schemaVt);
			// extended properties (app.xml)
			ns.AddNamespace("xp", ExcelPackage.schemaExtended);
			// custom properties
			ns.AddNamespace("ctp", ExcelPackage.schemaCustom);
			// core properties
			ns.AddNamespace("cp", ExcelPackage.schemaCore);
			// core property namespaces
			ns.AddNamespace("dc", ExcelPackage.schemaDc);
			ns.AddNamespace("dcterms", ExcelPackage.schemaDcTerms);
			ns.AddNamespace("dcmitype", ExcelPackage.schemaDcmiType);
			ns.AddNamespace("xsi", ExcelPackage.schemaXsi);
			ns.AddNamespace("x14", ExcelPackage.schemaMain2009);
			ns.AddNamespace("xm", ExcelPackage.schemaOfficeMain2006);
			return ns;
		}

		/// <summary>
		/// Close the internal stream
		/// </summary>
		private void CloseStream()
		{
			// Issue15252: Clear output buffer
			if (this.Stream != null)
			{
				this.Stream.Close();
				this.Stream.Dispose();
			}

			_stream = new MemoryStream();
		}

		private void Load(Stream input, Stream output, string password)
		{
			//Release some resources:
			if (this._package != null)
			{
				this._package.Close();
				this._package = null;
			}
			if (this._stream != null)
			{
				this._stream.Close();
				this._stream.Dispose();
				this._stream = null;
			}
			this._isExternalStream = true;
			if (input.Length == 0) // Template is blank, Construct new
			{
				this._stream = output;
				this.ConstructNewFile(password);
			}
			else
			{
				Stream ms;
				this._stream = output;
				if (password != null)
				{
#if !MONO
					Stream encrStream = new MemoryStream();
					ExcelPackage.CopyStream(input, ref encrStream);
					EncryptedPackageHandler eph = new EncryptedPackageHandler();
					this.Encryption.Password = password;
					ms = eph.DecryptPackage((MemoryStream)encrStream, this.Encryption);
#endif
#if MONO
                    throw new NotSupportedException("Encryption is not supported under Mono.");
#endif
				}
				else
				{
					ms = new MemoryStream();
					ExcelPackage.CopyStream(input, ref ms);
				}

				try
				{
					//this._package = Package.Open(this._stream, FileMode.Open, FileAccess.ReadWrite);
					this._package = new Packaging.ZipPackage(ms);
				}
				catch (Exception ex)
				{
#if !MONO
					EncryptedPackageHandler eph = new EncryptedPackageHandler();
					if (password == null && CompoundDocument.IsStorageILockBytes(CompoundDocument.GetLockbyte((MemoryStream)_stream)) == 0)
					{
						throw new Exception("Can not open the package. Package is an OLE compound document. If this is an encrypted package, please supply the password", ex);
					}
					else
					{
						throw;
					}
#endif
#if MONO
                    throw;
#endif
				}
			}
			//Clear the workbook so that it gets reinitialized next time
			this._workbook = null;
		}
		#endregion
	}
}