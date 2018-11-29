using System;
using System.Linq;
using System.Resources;

namespace OfficeOpenXml.Internationalization
{
	/// <summary>
	/// Contains string resources that can be configured with translations.
	/// </summary>
	public class StringResources
	{
		#region Properties
		// TODO: This property is for example only and should be deleted.
		[StringResource("The {0} will be replaced with a field name. Place accordingly in your translations.")]
		public string GrandTotalCaption
		{
			get
			{
				if (this.ResourceManager == null)
					return "{0} Grand Total";
				else
					return this.ResourceManager.GetString(nameof(GrandTotalCaption));
			}
		}

		// TODO: This property is for example only and should be deleted.
		[StringResource]
		public string GroupByCaption
		{
			get
			{
				if (this.ResourceManager == null)
					return "Group by";
				else
					return this.ResourceManager.GetString(nameof(GroupByCaption));
			}
		}

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
	}
}
