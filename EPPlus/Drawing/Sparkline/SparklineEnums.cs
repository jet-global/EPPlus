using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.Sparkline
{
    public enum SparklineType
    {
        Line, Column, Stacked
    }

    public enum SparklineAxisMinMax
    {
        Individual, Group, Custom
    }

    public enum DispBlanksAs
    {
        Span, Gap, Zero
    }
}
