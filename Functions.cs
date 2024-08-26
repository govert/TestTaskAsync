using System;
using System.Diagnostics;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace TestTaskAsync
{
    public class Functions
    {
        // This is the internal 'async' function
        static async Task<object> SlowAddImpl(double d1, double d2)
        {
            await Task.Delay(2000);
            return d1 + d2;
        }

        // This is the Excel UDF that wraps the internal 'async' function
        [ExcelFunction("is a slow function to add numbers")]
        public static object SlowAdd(double d1, double d2)
        {
            return ExcelTaskUtil.RunTask("SlowAdd", new object[] { d1, d2 }, () => SlowAddImpl(d1, d2));
        }

        // Upgrading to ExcelDna.AddIn v 1.90-alpha2 allow the function to loook like this, without needing any extra helpers
        //[ExcelFunction("is a slow function to add numbers")]
        //public static async Task<object> SlowAdd(double d1, double d2)
        //{
        //    await Task.Delay(2000);
        //    Debug.WriteLine($"Completed {d1} + {d2}");
        //    return d1 + d2;
        //}

    }
}
