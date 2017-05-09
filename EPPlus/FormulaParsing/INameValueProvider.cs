namespace OfficeOpenXml.FormulaParsing
{
	public interface INameValueProvider
	{
		bool IsNamedValue(string key, string worksheet);

		object GetNamedValue(string key);

		void Reload();
	}
}
