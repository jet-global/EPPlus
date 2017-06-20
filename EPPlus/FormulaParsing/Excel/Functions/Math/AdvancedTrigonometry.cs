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
 * Mats Alm   		                Added		                2015-01-11
 *******************************************************************************/
using MathObj = System.Math;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// Implements AdvancedTrionometry functions.
	/// Thanks to the guys in this thread: http://stackoverflow.com/questions/2840798/c-sharp-math-class-question
	/// </summary>
	public static class AdvancedTrigonometry
	{
		/// <summary>
		/// Handles the calculation for Hyperbolic Secant.
		/// </summary>
		/// <param name="x">A double to be evaluated.</param>
		/// <returns>Returns the hyperbolic secant of an angle.</returns>
		public static double HyperbolicSecant(double x)
		{
			return 2 / (MathObj.Exp(x) + MathObj.Exp(-x));
		}

		/// <summary>
		/// Handles the calculation for Secant.
		/// </summary>
		/// <param name="x">A double to be evaluated.</param>
		/// <returns>Returns the secant of an angle.</returns>
		public static double Secant(double x)
		{
			return 1 / MathObj.Cos(x);
		}

		/// <summary>
		/// Handles the calculation for Inverse Hyperbolic Tangent.
		/// </summary>
		/// <param name="x">A double to be evaluated.</param>
		/// <returns>Returns the inverse hyperbolic tangent of a number.</returns>
		public static double InverseHyperbolicTangent(double x)
		{
			return MathObj.Log((1 + x) / (1 - x)) / 2;
		}

		/// <summary>
		/// Handles the calculation for inverse hyperbolic sine.
		/// </summary>
		/// <param name="x">A double to be evaluated.</param>
		/// <returns>Returns the inverse hyperbolic sine of a number. </returns>
		public static double InverseHyperbolicSine(double x)
		{
			return MathObj.Log(x + MathObj.Sqrt(x * x + 1.0));
		}

		/// <summary>
		/// Handles the calculation for inverse hyperbolic cotangent.
		/// </summary>
		/// <param name="x">A double to be evaluated.</param>
		/// <returns>Returns the inverse hyperbolic cotangent of a number.</returns>
		public static double InverseHyperbolicCotangent(double x)
		{
			return MathObj.Log((x + 1.0) / (x - 1.0)) / 2.0;
		}

		/// <summary>
		/// Handles the calculation for inverse cotangent.
		/// </summary>
		/// <param name="x">A double to be evaluated.</param>
		/// <returns>Returns the principal value of the arccotangent, or inverse cotangent, of a number.</returns>
		public static double InverseCotangent(double x)
		{
			return 2.0 * MathObj.Atan(1) - MathObj.Atan(x);
		}

		/// <summary>
		/// Handles the calculation for inverse hyperbolic cosine.
		/// </summary>
		/// <param name="x">A double to be evaluated.</param>
		/// <returns>Returns the inverse hyperbolic cosine of a number.</returns>
		public static double InverseHyperbolicCosine(double x)
		{
			return MathObj.Log(x + MathObj.Sqrt(x * x - 1));
		}

		/// <summary>
		/// Handles the calculation for hyperbolic cosecant.
		/// </summary>
		/// <param name="x">A double to be evaluated.</param>
		/// <returns>Return the hyperbolic cosecant of an angle specified in radians.</returns>
		public static double HyperbolicCosecant(double x)
		{
			if ((MathObj.Exp(x) - MathObj.Exp(-x)) == 0)
				return -2;
			return 2 / (MathObj.Exp(x) - MathObj.Exp(-x));
		}

		/// <summary>
		/// Handles the calculation for cotangent.
		/// </summary>
		/// <param name="x">A double to be evaluated.</param>
		/// <returns>Return the cotangent of an angle specified in radians.</returns>
		public static double Cotangent(double x)
		{
			if (MathObj.Tan(x) == 0)
				return -2;
			return 1 / MathObj.Tan(x);
		}

		/// <summary>
		/// Handles the calculation for hyperbolic cotangent.
		/// </summary>
		/// <param name="x">A double to be evaluated.</param>
		/// <returns>Return the hyperbolic cotangent of a hyperbolic angle.</returns>
		public static double HyperbolicCotangent(double x)
		{
			//This is to handle a rounding diffrence between excel and EPPlus. 
			var NaNChecker = (MathObj.Exp(x) + MathObj.Exp(-x)) / (MathObj.Exp(x) - MathObj.Exp(-x));
			if (NaNChecker.Equals(double.NaN))
				return 1;

			return (MathObj.Exp(x) + MathObj.Exp(-x)) / (MathObj.Exp(x) - MathObj.Exp(-x));
		}

		/// <summary>
		/// Handles the calculation for cosecant.
		/// </summary>
		/// <param name="x">A double to be evaluated.</param>
		/// <returns>Returns the cosecant of an angle specified in radians.</returns>
		public static double Cosecant(double x)
		{
			return 1 / MathObj.Sin(x);
		}

		/// <summary>
		/// This is used as a check to make sure a divide by zero error is caught.
		/// </summary>
		/// <param name="x">The double to be evaluated.</param>
		/// <param name="SinValue">This is a out value. if true = -2 else it is Sin(x)</param>
		/// <returns>This returns true or false.</returns>
		public static bool TryCheckIfCosecantWillHaveADivideByZeroError(double x, out double SinValue)
		{
			if(MathObj.Sin(x) == 0)
			{
				SinValue = -2;
				return true;
			}
			SinValue = MathObj.Sin(x);
			return false;
		}

	}
}