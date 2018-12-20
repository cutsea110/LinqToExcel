using System;
using System.Linq;
using ClosedXML.Excel;
using Excel.Linq;

namespace Sample
{
    public class Program
    {
        static void Main(string[] args)
        {
            using (var ctx = new XlsWorkbook(@"C:\Users\cut-sea\project\SSU\force\src\Force\dgcommon\Data\Doc\system.xlsx"))
            {
                //var list = ctx.Where(s => !string.IsNullOrEmpty((string)s.Cell("C1").Value) && ((string)s.Cell("C1").Value).StartsWith("Student"));
                //var list2 = ctx.Where(s => s.Name.StartsWith("健康") || s.Name.Contains("学生"));
                var list = ctx.Where(s => !string.IsNullOrEmpty((string)s.Cell("C1").Value));
                foreach(var s in list)
                {
                    Console.WriteLine($"{s.Name}: {s.Cell("C1").Value} {s.RowsUsed().Count()}行");
                }
            }

        }
    }
}
