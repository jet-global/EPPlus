using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Sparkline;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EPPlusTest
{
    [TestClass]
    public class SparklineTests
    {
        #region Test Constants
        string rootNode = @"<ext uri=""{05C60535-1F16-4fd2-B633-F4F36F0B64E0
    }"" xmlns:x14=""http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"">

            <x14:sparklineGroups xmlns:xm=""http://schemas.microsoft.com/office/excel/2006/main"">
				<x14:sparklineGroup type=""stacked"" displayEmptyCellsAs=""gap"" negative=""1"">
					<x14:colorSeries rgb=""FF376092""/>
					<x14:colorNegative rgb=""FFD00000""/>
					<x14:colorAxis rgb=""FF000000""/>
					<x14:colorMarkers rgb=""FFD00000""/>
					<x14:colorFirst rgb=""FFD00000""/>
					<x14:colorLast rgb=""FFD00000""/>
					<x14:colorHigh rgb=""FFD00000""/>
					<x14:colorLow rgb=""FFD00000""/>
					<x14:sparklines>
						<x14:sparkline>
							<xm:f>Sheet1!F7:F7</xm:f>
							<xm:sqref>G7</xm:sqref>
						</x14:sparkline>
						<x14:sparkline>
							<xm:f>Sheet1!F8:F8</xm:f>
							<xm:sqref>G8</xm:sqref>
						</x14:sparkline>
						<x14:sparkline>
							<xm:f>Sheet1!F9:F9</xm:f>
							<xm:sqref>G9</xm:sqref>
						</x14:sparkline>
						<x14:sparkline>
							<xm:f>Sheet1!F10:F10</xm:f>
							<xm:sqref>G10</xm:sqref>
						</x14:sparkline>
						<x14:sparkline>
							<xm:f>Sheet1!F11:F11</xm:f>
							<xm:sqref>G11</xm:sqref>
						</x14:sparkline>
						<x14:sparkline>
							<xm:f>Sheet1!F12:F12</xm:f>
							<xm:sqref>G12</xm:sqref>
						</x14:sparkline>
						<x14:sparkline>
							<xm:f>Sheet1!F13:F13</xm:f>
							<xm:sqref>G13</xm:sqref>
						</x14:sparkline>
						<x14:sparkline>
							<xm:f>Sheet1!F14:F14</xm:f>
							<xm:sqref>G14</xm:sqref>
						</x14:sparkline>
						<x14:sparkline>
							<xm:f>Sheet1!F15:F15</xm:f>
							<xm:sqref>G15</xm:sqref>
						</x14:sparkline>
						<x14:sparkline>
							<xm:f>Sheet1!F16:F16</xm:f>
							<xm:sqref>G16</xm:sqref>
						</x14:sparkline>
					</x14:sparklines>
				</x14:sparklineGroup>
				<x14:sparklineGroup type=""stacked"" displayEmptyCellsAs=""gap"" negative=""1"">
					<x14:colorSeries rgb=""FF376092""/>
					<x14:colorNegative rgb=""FFD00000""/>
					<x14:colorAxis rgb=""FF000000""/>
					<x14:colorMarkers rgb=""FFD00000""/>
					<x14:colorFirst rgb=""FFD00000""/>
					<x14:colorLast rgb=""FFD00000""/>
					<x14:colorHigh rgb=""FFD00000""/>
					<x14:colorLow rgb=""FFD00000""/>
					<x14:sparklines>
						<x14:sparkline>
							<xm:f>Sheet1!F6:F6</xm:f>
							<xm:sqref>G6</xm:sqref>
						</x14:sparkline>
						<x14:sparkline>
							<xm:f>Sheet1!F17:F17</xm:f>
							<xm:sqref>G17</xm:sqref>
						</x14:sparkline>
					</x14:sparklines>
				</x14:sparklineGroup>
			</x14:sparklineGroups>
        </ext>";
        #endregion

        [TestMethod]
        public void ParseExcelSparklineGroups()
        {
            XmlDocument extensions = new XmlDocument();
            extensions.LoadXml(this.rootNode);
            var sparklineGroupsXml = extensions.ChildNodes[0].ChildNodes[0];
            Assert.AreEqual(2, sparklineGroupsXml.ChildNodes.Count);
            XmlNamespaceManager manager = new XmlNamespaceManager(extensions.NameTable);
            // TODO: Figure out why I have to do this
            manager.AddNamespace("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");

            var sparklineGroups = new ExcelSparklineGroups(manager, sparklineGroupsXml);
            Assert.AreEqual(2, sparklineGroups.SparklineGroups.Count);
            var group1 = sparklineGroups.SparklineGroups[0];
            // Test that unused optional attributes are not set.
            Assert.IsNull(group1.ManualMin);
            Assert.IsNull(group1.ManualMax);
            Assert.IsNull(group1.LineWeight);
            Assert.IsFalse(group1.DateAxis);
            Assert.IsFalse(group1.Markers);
            Assert.IsFalse(group1.High);
            Assert.IsFalse(group1.Low);
            Assert.IsFalse(group1.First);
            Assert.IsFalse(group1.Last);
            Assert.IsFalse(group1.DisplayXAxis);
            Assert.IsFalse(group1.DisplayYAxis);
            Assert.AreEqual(SparklineAxisMinMax.Individual, group1.MinAxisType);
            Assert.AreEqual(SparklineAxisMinMax.Individual, group1.MaxAxisType);
            Assert.IsFalse(group1.RightToLeft);

            // Test that attributes are parsed correctly.
            Assert.AreEqual(SparklineType.Stacked, group1.Type);
            Assert.AreEqual(DispBlanksAs.Gap, group1.DisplayEmptyCellsAs);
            Assert.AreEqual(true, group1.Negative);

            // Test that color subnodes are parsed correctly.
            var darkRed = Color.FromArgb(unchecked((int)0xFFD00000));
            Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF376092)), sparklineGroups.SparklineGroups[0].ColorSeries);
            Assert.AreEqual(darkRed, group1.ColorNegative);
            Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF000000)), group1.ColorAxis);
            Assert.AreEqual(darkRed, group1.ColorMarkers);
            Assert.AreEqual(darkRed, group1.ColorFirst);
            Assert.AreEqual(darkRed, group1.ColorLast);
            Assert.AreEqual(darkRed, group1.ColorHigh);
            Assert.AreEqual(darkRed, group1.ColorLow);
            
        }
    }
}
