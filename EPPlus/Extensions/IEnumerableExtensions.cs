using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Extensions
{
	/// <summary>
	/// Represents a class which contains extension methods for the <see cref="System.Collections.Generic.IEnumerable{T}"/> object.
	/// </summary>
	public static class IEnumerableExtensions
	{
		/// <summary>
		/// Performs the specified <paramref name="action"/> on each element of the <paramref name="list"/>.
		/// </summary>
		/// <typeparam name="T">The type of object contained in the <paramref name="list"/>.</typeparam>
		/// <param name="list">The list to enumerate.</param>
		/// <param name="action">The action to perform.</param>
		/// <exception cref="System.ArgumentNullException">Thrown when <paramref name="action"/> or <paramref name="list"/> is null.</exception>
		public static void ForEach<T>(this IEnumerable<T> list, Action<T> action)
		{
			if (action == null)
				throw new ArgumentNullException("action");
			if (list == null)
				throw new ArgumentNullException("list");
			foreach (T element in list)
				action(element);
		}

		/// <summary>
		/// Performs the specified <paramref name="action"/> on each element of the <paramref name="list"/>.
		/// </summary>
		/// <typeparam name="T">The type of object contained in the <paramref name="list"/>.</typeparam>
		/// <param name="list">The list to enumerate.</param>
		/// <param name="action">The action to execute and the index of the source element.</param>
		/// <exception cref="System.ArgumentNullException">Thrown when <paramref name="action"/> or <paramref name="list"/> is null.</exception>
		public static void ForEach<T>(this IEnumerable<T> list, Action<T, int> action)
		{
			if (action == null)
				throw new ArgumentNullException("action");
			if (list == null)
				throw new ArgumentNullException("list");
			int i = 0;
			foreach (T element in list)
				action(element, i++);
		}
	}
}
