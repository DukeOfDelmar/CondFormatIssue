using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using ClosedXML.Report;

namespace CondFormatIssue
{
    public class FruitSales
    {
        public List<FruitSalesDetail> details { get; set; }
    }

    public class FruitSalesDetail
    {
        public string name;
        public int lastMonth;
        public int thisMonth;
    }

    class Program
    {
        static void Main(string[] args)
        {
            var reportData = new FruitSales
            {
                details = new List<FruitSalesDetail>
                {
                    new FruitSalesDetail { name = "Apple", lastMonth = 14, thisMonth = 17 },
                    new FruitSalesDetail { name = "Banana", lastMonth = 18, thisMonth = 15 },
                    new FruitSalesDetail { name = "Cherry", lastMonth = 207, thisMonth = 142 },
                }
            };

            GenerateTemplate(reportData, "Working.xlsx"); // These two workbooks are identical, except "Broken.xlsx" applies
            GenerateTemplate(reportData, "Broken.xlsx"); // conditional formatting to one of the cells in the vertical table.
        }

        private static void GenerateTemplate(FruitSales reportData, string templateName)
        {
            var templatePath = Path.Combine(GetApplicationPath(), "Templates", templateName);
            var template = new XLTemplate(templatePath);
            template.AddVariable(reportData);
            template.Generate();
        }

        public static string GetApplicationPath()
        {
            return new Uri(Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase)).LocalPath;
        }

    }
}
