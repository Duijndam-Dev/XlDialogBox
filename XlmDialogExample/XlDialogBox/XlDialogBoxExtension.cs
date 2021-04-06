using System;
using ExcelDna.Integration;

namespace ExcelDna.XlDialogBox
{

    // do more than checking for a null-pointer to see if an object under XlDialogBox is valid
    internal static class Extensions
    {
        public static bool IsNull(this object instance)
        {
            return 
                instance == null || 
                instance == System.Type.Missing ||
                instance is DBNull ||
                instance is ExcelEmpty ||
                instance is ExcelError ||
                instance is ExcelMissing;
        }
    }
}