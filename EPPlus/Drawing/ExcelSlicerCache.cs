using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
	public class ExcelSlicerCache: XmlHelper
	{
		#region Constructors
		internal ExcelSlicerCache(XmlNode node, XmlNamespaceManager namespaceManager): base(namespaceManager, node)
		{

		}
		#endregion
	}
}
