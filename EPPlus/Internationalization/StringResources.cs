using System;
using System.Linq;
using System.Resources;
using System.Runtime.CompilerServices;

namespace OfficeOpenXml.Internationalization
{
	/// <summary>
	/// Contains string resources that can be configured with translations.
	/// </summary>
	public class StringResources
	{
		#region Properties
		[StringResource("The {0} will be replaced with a field name. Place accordingly in your translations.")]
		public string TotalCaptionWitFollowingValue => this.GetValue("Total {0}");

		[StringResource("The {0} will be replaced with a field name. Place accordingly in your translations.")]
		public string TotalCaptionWitPrecedingValue => this.GetValue("{0} Total");

		[StringResource]
		public string GrandTotalCaption => this.GetValue("Grand Total");

		private ResourceManager ResourceManager { get; set; }
		#endregion

		#region Public Methods
		/// <summary>
		/// Loads a <see cref="ResourceManager"/> to retrieve translated string resources from.
		/// </summary>
		/// <param name="manager">The <see cref="ResourceManager"/> to load.</param>
		public void LoadResourceManager(ResourceManager manager)
		{
			if (manager == null)
				throw new ArgumentNullException(nameof(manager));
			this.ResourceManager = manager;
		}

		/// <summary>
		/// Validates that a <see cref="ResourceManager"/> has been loaded and contains all of the 
		/// necessary string resources.
		/// </summary>
		/// <param name="error">The error that occurred, if any.</param>
		/// <returns>True if the <see cref="ResourceManager"/> has been loaded and contains the necessary keys, otherwise false.</returns>
		public bool ValidateLoadedResourceManager(out string error)
		{
			error = null;
			if (this.ResourceManager == null)
			{
				error = "No resource manager loaded.";
				return false;
			}
			var resourceKeys = this.GetType()
				.GetProperties(System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.Public)
				.Where(p => p.GetCustomAttributesData().Any(a => a.AttributeType == typeof(StringResource)))
				.Select(p => p.Name);
			var missingStringResources = resourceKeys.Where(k => this.ResourceManager.GetString(k) == null);
			if (missingStringResources.Any())
			{
				error = $"The following string resources were missing:{Environment.NewLine}{string.Join(", ", missingStringResources)}";
				return false;
			}
			return true;
		}
		#endregion

		#region Private Methods
		private string GetValue(string defaultValue, [CallerMemberName] string memberName = "")
		{
			if (this.ResourceManager == null)
				return defaultValue;
			return this.ResourceManager.GetString(memberName);
		}
		#endregion
	}
}
