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
 * Mats Alm   		                Added		                2013-12-03
 * Eric Beiler                      Enable Multiple Selections  2015-09-01
 *******************************************************************************/
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
	public class Choose : ExcelFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var items = new List<object>();
			for (int x = 0; x < arguments.Count(); x++)
			{
				items.Add(arguments.ElementAt(x).ValueFirst);
			}

			var chooseIndeces = arguments.ElementAt(0).ValueFirst as IEnumerable<FunctionArgument>;
			if (chooseIndeces != null && chooseIndeces.Count() > 1)
			{
				IntArgumentParser intParser = new IntArgumentParser();
				object[] values = chooseIndeces.Select(chosenIndex => items[(int)intParser.Parse(chosenIndex.ValueFirst)]).ToArray();
				return CreateResult(values, DataType.Enumerable);
			}
			else
			{
				var index = ArgToInt(arguments, 0);
				return CreateResult(items[index].ToString(), DataType.String);
			}
		}
	}
}
