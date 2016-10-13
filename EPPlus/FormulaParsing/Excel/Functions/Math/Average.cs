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
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class Average : HiddenValuesHandlingFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            if(ValidateArguments(arguments, 1) == false)
				return new CompileResult(eErrorType.Div0);
            double nValues = 0d, result = 0d;
            foreach (var arg in arguments)
            {
                var error = Calculate(arg, context, ref result, ref nValues);
				if (error != null)
					return new CompileResult(error.Value);
            }
            return CreateResult(Divide(result, nValues), DataType.Decimal);
        }

        private eErrorType? Calculate(FunctionArgument arg, ParsingContext context, ref double retVal, ref double nValues, bool isInArray = false)
        {
            if (ShouldIgnore(arg))
            {
                return null;
            }
            if (arg.Value is IEnumerable<FunctionArgument>)
            {
                foreach (var item in (IEnumerable<FunctionArgument>)arg.Value)
                {
                    var error = Calculate(item, context, ref retVal, ref nValues, true);
					if (error != null)
						return error;
                }
            }
            else if (arg.IsExcelRange)
            {
                foreach (var c in arg.ValueAsRangeInfo)
                {
                    if (ShouldIgnore(c, context)) continue;
                    CheckForAndHandleExcelError(c);
                    if (!IsNumeric(c.Value) || c.Value is bool) continue;
                    nValues++;
                    retVal += c.ValueDouble;
                }
            }
            else
            {
                var numericValue = GetNumericValue(arg.Value, isInArray);
				if (numericValue.HasValue)
				{
					nValues++;
					retVal += numericValue.Value;
				}
				else if (arg.Value is string && !isInArray)
				{
					return eErrorType.Value;
				}
            }
            CheckForAndHandleExcelError(arg);
			return null;
        }

        private double? GetNumericValue(object obj, bool isInArray)
        {
            if (IsNumeric(obj) && !(obj is bool))
            {
                return ConvertUtil.GetValueDouble(obj);
            }
			if (!isInArray)
			{
				double number;
				System.DateTime date;
				if (obj is bool)
				{
					return ConvertUtil.GetValueDouble(obj);
				}
				else if (ConvertUtil.TryParseNumericString(obj, out number))
				{
					return number;
				}
				else if (ConvertUtil.TryParseDateString(obj, out date))
				{
					return date.ToOADate();
				}
			}
            return default(double?);
        }
    }
}
