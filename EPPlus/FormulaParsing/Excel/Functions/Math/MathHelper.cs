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
	/// Thanks to the guys in this thread: http://stackoverflow.com/questions/2840798/c-sharp-math-class-question
	/// </summary>
	public static class MathHelper
	{
		// Secant 
		public static double Sec(double x)
		{
			return 1 / MathObj.Cos(x);
		}

		// Cosecant
		public static double Cosec(double x)
		{
			if (MathObj.Sin(x) == 0)
				return -2;
			return 1 / MathObj.Sin(x);
		}

		// Cotangent 
		public static double Cotan(double x)
		{
			if (MathObj.Tan(x) == 0)
				return -2;
			return 1 / MathObj.Tan(x);
		}

		// Inverse Cotangent 
		public static double Arccotan(double x)
		{
			return 2 * MathObj.Atan(1) - MathObj.Atan(x);
		}

		// Hyperbolic Secant 
		public static double HSec(double x)
		{
			return 2 / (MathObj.Exp(x) + MathObj.Exp(-x));
		}

		// Hyperbolic Cosecant 
		public static double HCosec(double x)
		{
			if ((MathObj.Exp(x) - MathObj.Exp(-x)) == 0)
				return -2;
			return 2 / (MathObj.Exp(x) - MathObj.Exp(-x));
		}

		// Hyperbolic Cotangent 
		public static double HCotan(double x)
		{
			if ((MathObj.Exp(x) - MathObj.Exp(-x) == 0))
				return -2;
			var NaNChecker = (MathObj.Exp(x) + MathObj.Exp(-x)) / (MathObj.Exp(x) - MathObj.Exp(-x));
			if (NaNChecker.Equals(double.NaN))
				return 1;

			return (MathObj.Exp(x) + MathObj.Exp(-x)) / (MathObj.Exp(x) - MathObj.Exp(-x));
		}

		// Inverse Hyperbolic Sine 
		public static double HArcsin(double x)
		{
			return MathObj.Log(x + MathObj.Sqrt(x * x + 1));
		}

		// Inverse Hyperbolic Cosine 
		public static double HArccos(double x)
		{
			return MathObj.Log(x + MathObj.Sqrt(x * x - 1));
		}

		// Inverse Hyperbolic Tangent 
		public static double HArctan(double x)
		{
			return MathObj.Log((1 + x) / (1 - x)) / 2;
		}

		public static double HArccotan(double x)
		{
			return MathObj.Log((x + 1) / (x - 1)) / 2;
		}
	}
}
