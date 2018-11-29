using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Internationalization;

namespace EPPlusTest.Internationalization
{
	[TestClass]
	public class StringResourcesTest
	{
		#region LoadResourceManager Tests
		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void LoadResourceManagerNullResourceManagerThrowsException()
		{
			var stringResources = new StringResources();
			stringResources.LoadResourceManager(null);
		}
		#endregion

		#region ValidateLoadedResourceManager Tests
		[TestMethod]
		public void ValidateLoadedResourceManagerNullManagerReturnsFalse()
		{
			var stringResources = new StringResources();
			var result = stringResources.ValidateLoadedResourceManager(out string error);
			Assert.IsNotNull(error);
			Assert.IsFalse(result);
		}
		#endregion
	}
}
