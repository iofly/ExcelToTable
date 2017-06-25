using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using System.Security.Cryptography;
using System.Text.RegularExpressions;

namespace ExcelToTable
{
    public enum ResultCode { Success, ErrorOpeningFile };

    class Program
	{
        static void cMain(string[] args)
        {
            var r = ExcelReader.ParseExcelRange("AB2:AC3");
            DateTime dt = new DateTime();
        }

        static void Main(string[] args)
        {
            ResultCode rc;
            string ResultDesc = string.Empty;
            List<List<string>> rows;
            string fileext = ".txt";
            string text = string.Empty;

            #region Define supported arguments
            List<SimpleArg> supportedArgs = new List<SimpleArg>();
            supportedArgs.Add(new SimpleArg { Name = "-filename", IsSwitch = false, Required = true, DefaultValue = null, ArgType = SimpleArgType.String });
            supportedArgs.Add(new SimpleArg { Name = "-format", IsSwitch = false, Required = true, DefaultValue = "html", ArgType = SimpleArgType.String });
            supportedArgs.Add(new SimpleArg { Name = "-worksheet", IsSwitch = false, Required = false, DefaultValue = 1, ArgType = SimpleArgType.Integer });
            supportedArgs.Add(new SimpleArg { Name = "-range", IsSwitch = false, Required = false, DefaultValue = null, ArgType = SimpleArgType.String });
            #endregion

            #region Parse and validate arguments
            
            SimpleArgParser parser = new SimpleArgParser(supportedArgs);
            Dictionary<string, dynamic> ar = new Dictionary<string, dynamic>();

            //Validate against general rules as specified above.
            try
            {
                ar = parser.ParseArgs(args);
                parser.ValidateArgs(ar);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Utils.ShowUsage();
                return;
            }

            //Fix for filename ->abs filename - for case where app is called via batch file
            if (!System.IO.Path.IsPathRooted(ar["-filename"]))
            {
                ar["-filename"] = System.IO.Path.GetFullPath(ar["-filename"]);
            }


            WorkSheetRangeCoordinates wsrc = null;
            if (ar.ContainsKey("-range"))
            {
                wsrc = ExcelReader.ParseExcelRange(ar["-range"]);
            }

            //App specific validation of args
            try
            {
                Utils.ValidateAppArgs(ar);
            }
            catch(ArgumentException ex)
            {
                Console.WriteLine(ex.Message);
                Utils.ShowUsage();
                return;
            }
            #endregion



            




            #region Read excel file
            rows = ExcelReader.ReadExcelRows(ar["-filename"], out rc, out ResultDesc, ar["-worksheet"], wsrc);
            if (rc == ResultCode.ErrorOpeningFile)
            {
                Console.WriteLine(String.Format("ErrorOpeningFile: {0}", ResultDesc));
                return;
            }
            #endregion

            #region Output result
            if (ar["-format"] == "wikitable")
            {
                text = Utils.RowsToWikiTable(rows);
                fileext = ".txt";
            }
            else if (ar["-format"] == "jsonsobjects")
            {
                text = Utils.RowsToJSON_ArrayOfObjects(rows);
                fileext = ".objects.json";
            }
            else if (ar["-format"] == "jsonarrays")
            {
                text = Utils.RowsToJSON_ArrayOfArrays(rows);
                fileext = ".arrays.json";
            }
            else
            {
                text = Utils.RowsToHTMLTable(rows);
                fileext = ".html";
            }

            

            System.IO.File.WriteAllText(System.IO.Path.GetFileName(ar["-filename"]) + fileext, text);
            #endregion
        }
    }
}
