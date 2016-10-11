using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicers
{
	/// <summary>
	/// Represents a single Excel Slicer in an <see cref="ExcelSlicers"/> file.
	/// </summary>
	public class ExcelSlicer: XmlHelper
	{
		#region Class Variables
		private string _name;
		private static XmlNamespaceManager _slicerDocumentNamespaceManager;
		#endregion

		#region Properties
		/// <summary>
		/// Gets or sets the slicer cache associated with this slicer.
		/// </summary>
		public ExcelSlicerCache SlicerCache { get; set; }

		/// <summary>
		/// Gets or sets this slicer's name attribute.
		/// </summary>
		public string Name
		{
			get
			{
				if (this._name == null)
					this._name = this.TopNode.Attributes["name"].Value;
				return this._name;
			}
			set
			{
				this._name = value;
				this.TopNode.Attributes["name"].Value = value;
			}
		}

		/// <summary>
		/// Gets a namespace manager that contains the namespaces that are used by slicer.xml files.
		/// </summary>
		public static XmlNamespaceManager SlicerDocumentNamespaceManager
		{
			get
			{
				if (ExcelSlicer._slicerDocumentNamespaceManager == null)
				{
					var nameTable = new NameTable();
					var namespaceManager = new XmlNamespaceManager(nameTable);
					namespaceManager.AddNamespace(string.Empty, ExcelPackage.schemaMain2009);
					// Hack to work around a bug where SelectSingleNode ignores the default namespace.
					namespaceManager.AddNamespace("default", ExcelPackage.schemaMain2009);
					namespaceManager.AddNamespace("mc", ExcelPackage.schemaMarkupCompatibility);
					_slicerDocumentNamespaceManager = namespaceManager;
				}
				return ExcelSlicer._slicerDocumentNamespaceManager;
			}
		}

		private ExcelWorksheet Worksheet { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Initialize a new <see cref="ExcelSlicer"/> based on the specified <paramref name="node"/>.
		/// </summary>
		/// <param name="node">The "<slicer/>" node that this slicer represents.</param>
		/// <param name="namespaceManager">The namespace manager to use to process the slicer.</param>
		/// <param name="worksheet">The <see cref="ExcelWorksheet"/> that the slicer's drawing exists on.</param>
		internal ExcelSlicer(XmlNode node, XmlNamespaceManager namespaceManager, ExcelWorksheet worksheet): base(namespaceManager, node)
		{
			this.Worksheet = worksheet;
			var cacheName = node.Attributes["cache"].Value;
			this.SlicerCache = this.Worksheet.Workbook.SlicerCaches.Last(cache => cache.Name == cacheName);
			this.SlicerCache.Slicer = this;
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Updates this slicer's name and cache name variables to be new, unique values. 
		/// </summary>
		internal void IncrementNameAndCacheName()
		{
			var newSlicerNumber = this.Worksheet.Workbook.NextSlicerIdNumber[this.Name]++;
			this.SlicerCache.Name += newSlicerNumber.ToString();
			this.Worksheet.Workbook.Names.AddFormula(this.SlicerCache.Name, "#N/A");
			this.SlicerCache.TabId = this.Worksheet.SheetID.ToString();
			this.Name += $" {newSlicerNumber}";
		}
		#endregion
	}
}
