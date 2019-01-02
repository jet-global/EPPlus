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
		public void LoadResourceManagerNullResourceManager()
		{
			var stringResources = new StringResources();
			// Verify that setting the resource manager to null does not throw an exception.
			stringResources.LoadResourceManager(null);
		}

		[TestMethod]
		public void LoadResourceManagerSucceeds()
		{
			var stringResources = new StringResources();
			stringResources.LoadResourceManager(TestInternationalizationResources.ResourceManager);
			Assert.AreEqual("Different Total {0}", stringResources.TotalCaptionWithFollowingValue);
			Assert.AreEqual("{0} Default Total", stringResources.TotalCaptionWithPrecedingValue);
			Assert.AreEqual("Grand Total Sum", stringResources.GrandTotalCaption);
		}
		#endregion

		#region ValidateLoadedResourceManager Tests
		[TestMethod]
		public void ValidateLoadedResourceManagerSucceeds()
		{
			var stringResources = new StringResources();
			stringResources.LoadResourceManager(TestInternationalizationResources.ResourceManager);
			var result = stringResources.ValidateLoadedResourceManager(out string error);
			Assert.IsTrue(result);
			Assert.IsNull(error);
		}

		[TestMethod]
		public void ValidateLoadedResourceManagerNullManagerReturnsFalse()
		{
			var stringResources = new StringResources();
			var result = stringResources.ValidateLoadedResourceManager(out string error);
			Assert.IsNotNull(error);
			Assert.IsFalse(result);
		}

		[TestMethod]
		public void ValidateLoadedResourceManagerMissingKeyReturnsFalse()
		{
			var stringResources = new StringResources();
			stringResources.LoadResourceManager(TestInternationalizationResourcesMissingValue.ResourceManager);
			var result = stringResources.ValidateLoadedResourceManager(out string error);
			Assert.IsFalse(result);
			var expectedError = $"The following string resources were missing:{Environment.NewLine}{nameof(stringResources.TotalCaptionWithPrecedingValue)}";
			Assert.AreEqual(expectedError, error);
		}
		#endregion
	}
}
