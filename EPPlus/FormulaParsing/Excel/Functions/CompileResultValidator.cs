/* Copyright (C) 2011  Jan Källman
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
 *******************************************************************************
 * Mats Alm   		                Added		                2013-12-26
 *******************************************************************************/

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
	#region CompileResultValidator Abstract Methods
	public abstract class CompileResultValidator
	{
		/// <summary>
		/// Validates the answer as an excel format.
		/// </summary>
		public abstract void Validate(object obj);

		/// <summary>
		/// Checks for a validation error.
		/// </summary>
		public abstract bool TryGetValidationError(object obj, out eErrorType error);

		private static CompileResultValidator _empty;

		public static CompileResultValidator Empty
		{
			get { return _empty ?? (_empty = new EmptyCompileResultValidator()); }
		}
	}
	#endregion

	internal class EmptyCompileResultValidator : CompileResultValidator
	{
		public override void Validate(object obj)
		{
			// empty validator - do nothing
		}

		public override bool TryGetValidationError(object obj, out eErrorType error)
		{
			error = eErrorType.Null;
			return true;
		}
	}

}
