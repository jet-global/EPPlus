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
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Numeric;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
	public class BuiltInFunctions : FunctionsModule
	{
		public BuiltInFunctions()
		{
			// Text
			this.Functions["len"] = new Len();
			this.Functions["lower"] = new Lower();
			this.Functions["upper"] = new Upper();
			this.Functions["left"] = new Left();
			this.Functions["right"] = new Right();
			this.Functions["mid"] = new Mid();
			this.Functions["replace"] = new Replace();
			this.Functions["rept"] = new Rept();
			this.Functions["substitute"] = new Substitute();
			this.Functions["concatenate"] = new Concatenate();
			this.Functions["char"] = new CharFunction();
			this.Functions["exact"] = new Exact();
			this.Functions["find"] = new Find();
			this.Functions["fixed"] = new Fixed();
			this.Functions["proper"] = new Proper();
			this.Functions["search"] = new Search();
			this.Functions["text"] = new Text.Text();
			this.Functions["t"] = new T();
			this.Functions["hyperlink"] = new Hyperlink();
			this.Functions["value"] = new Value();
			// Numbers
			this.Functions["int"] = new IntFunction();
			// Math
			this.Functions["abs"] = new Abs();
			this.Functions["asin"] = new Asin();
			this.Functions["asinh"] = new Asinh();
			this.Functions["cos"] = new Cos();
			this.Functions["cosh"] = new Cosh();
			this.Functions["power"] = new Power();
			this.Functions["sign"] = new Sign();
			this.Functions["sqrt"] = new Sqrt();
			this.Functions["sqrtpi"] = new SqrtPi();
			this.Functions["pi"] = new Pi();
			this.Functions["product"] = new Product();
			this.Functions["ceiling"] = new Ceiling();
			this.Functions["count"] = new Count();
			this.Functions["counta"] = new CountA();
			this.Functions["countblank"] = new CountBlank();
			this.Functions["countif"] = new CountIf();
			this.Functions["countifs"] = new CountIfs();
			this.Functions["fact"] = new Fact();
			this.Functions["factdouble"] = new FactDouble();
			this.Functions["floor"] = new Floor();
			this.Functions["sin"] = new Sin();
			this.Functions["sinh"] = new Sinh();
			this.Functions["sum"] = new Sum();
			this.Functions["sumif"] = new SumIf();
			this.Functions["sumifs"] = new SumIfs();
			this.Functions["sumproduct"] = new SumProduct();
			this.Functions["sumsq"] = new Sumsq();
			this.Functions["stdev"] = new Stdev();
			this.Functions["stdevp"] = new StdevP();
			this.Functions["stdev.s"] = new StdevS();
			this.Functions["stdev.p"] = new StdevP();
			this.Functions["Stdevpa"] = new Stdevpa();
			this.Functions["Stdeva"] = new Stdeva();
			this.Functions["subtotal"] = new Subtotal();
			this.Functions["exp"] = new Exp();
			this.Functions["log"] = new Log();
			this.Functions["log10"] = new Log10();
			this.Functions["ln"] = new Ln();
			this.Functions["max"] = new Max();
			this.Functions["maxa"] = new Maxa();
			this.Functions["median"] = new Median();
			this.Functions["min"] = new Min();
			this.Functions["mina"] = new Mina();
			this.Functions["mod"] = new Mod();
			this.Functions["average"] = new Average();
			this.Functions["averagea"] = new AverageA();
			this.Functions["averageif"] = new AverageIf();
			this.Functions["averageifs"] = new AverageIfs();
			this.Functions["round"] = new Round();
			this.Functions["rounddown"] = new Rounddown();
			this.Functions["roundup"] = new Roundup();
			this.Functions["rand"] = new Rand();
			this.Functions["randbetween"] = new RandBetween();
			this.Functions["rank"] = new Rank();
			this.Functions["rank.eq"] = new Rank();
			this.Functions["rank.avg"] = new Rank(true);
			this.Functions["quotient"] = new Quotient();
			this.Functions["trunc"] = new Trunc();
			this.Functions["tan"] = new Tan();
			this.Functions["tanh"] = new Tanh();
			this.Functions["atan"] = new Atan();
			this.Functions["atan2"] = new Atan2();
			this.Functions["atanh"] = new Atanh();
			this.Functions["acos"] = new Acos();
			this.Functions["acosh"] = new Acosh();
			this.Functions["var"] = new Var();
			this.Functions["varp"] = new VarP();
			this.Functions["var.p"] = new VarP();
			this.Functions["var.s"] = new VarS();
			this.Functions["vara"] = new Vara();
			this.Functions["varpa"] = new Varpa();
			this.Functions["large"] = new Large();
			this.Functions["small"] = new Small();
			this.Functions["degrees"] = new Degrees();
			this.Functions["radians"] = new Radians();
			this.Functions["sec"] = new Sec();
			this.Functions["sech"] = new Sech();
			this.Functions["csc"] = new Csc();
			this.Functions["csch"] = new Csch();
			this.Functions["cot"] = new Cot();
			this.Functions["coth"] = new Coth();
			this.Functions["acot"] = new Acot();
			this.Functions["acoth"] = new Acoth();
			this.Functions["floor.math"] = new FloorMath();
			this.Functions["ceiling.math"] = new CeilingMath();
			// Information
			this.Functions["isblank"] = new IsBlank();
			this.Functions["isnumber"] = new IsNumber();
			this.Functions["istext"] = new IsText();
			this.Functions["isnontext"] = new IsNonText();
			this.Functions["iserror"] = new IsError();
			this.Functions["iserr"] = new IsErr();
			this.Functions["error.type"] = new ErrorType();
			this.Functions["iseven"] = new IsEven();
			this.Functions["isodd"] = new IsOdd();
			this.Functions["islogical"] = new IsLogical();
			this.Functions["isna"] = new IsNa();
			this.Functions["na"] = new Na();
			this.Functions["n"] = new N();
			// Logical
			this.Functions["if"] = new If();
			this.Functions["iferror"] = new IfError();
			this.Functions["ifna"] = new IfNa();
			this.Functions["not"] = new Not();
			this.Functions["and"] = new And();
			this.Functions["or"] = new Or();
			this.Functions["true"] = new True();
			this.Functions["false"] = new False();
			// Reference and lookup
			this.Functions["address"] = new Address();
			this.Functions["hlookup"] = new HLookup();
			this.Functions["vlookup"] = new VLookup();
			this.Functions["lookup"] = new Lookup();
			this.Functions["match"] = new Match();
			this.Functions["row"] = new Row();
			this.Functions["rows"] = new Rows();
			this.Functions["column"] = new Column();
			this.Functions["columns"] = new Columns();
			this.Functions["choose"] = new Choose();
			this.Functions["index"] = new Index();
			this.Functions["indirect"] = new Indirect();
			this.Functions["indirectaddress"] = new IndirectAddress();
			this.Functions["offset"] = new Offset();
			this.Functions["offsetaddress"] = new OffsetAddress();
			this.Functions["getpivotdata"] = new GetPivotData();
			// Date
			this.Functions["date"] = new Date();
			this.Functions["today"] = new Today();
			this.Functions["now"] = new Now();
			this.Functions["day"] = new Day();
			this.Functions["month"] = new Month();
			this.Functions["year"] = new Year();
			this.Functions["time"] = new Time();
			this.Functions["hour"] = new Hour();
			this.Functions["minute"] = new Minute();
			this.Functions["second"] = new Second();
			this.Functions["weeknum"] = new Weeknum();
			this.Functions["weekday"] = new Weekday();
			this.Functions["days360"] = new Days360();
			this.Functions["yearfrac"] = new Yearfrac();
			this.Functions["edate"] = new Edate();
			this.Functions["eomonth"] = new Eomonth();
			this.Functions["isoweeknum"] = new IsoWeekNum();
			this.Functions["workday"] = new Workday();
			this.Functions["workday.intl"] = new WorkdayIntl();
			this.Functions["networkdays"] = new Networkdays();
			this.Functions["networkdays.intl"] = new NetworkdaysIntl();
			this.Functions["datevalue"] = new DateValue();
			this.Functions["timevalue"] = new TimeValue();
			// Database
			this.Functions["dget"] = new Dget();
			this.Functions["dcount"] = new Dcount();
			this.Functions["dcounta"] = new DcountA();
			this.Functions["dmax"] = new Dmax();
			this.Functions["dmin"] = new Dmin();
			this.Functions["dsum"] = new Dsum();
			this.Functions["daverage"] = new Daverage();
			this.Functions["dvar"] = new Dvar();
			this.Functions["dvarp"] = new Dvarp();
		}
	}
}
