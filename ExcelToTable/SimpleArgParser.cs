using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;

namespace ExcelToTable
{
    public class SimpleArgParser
    {
        private List<SimpleArg> _SupportedArgs;
        private List<SimpleArg> _RequiredArgs;
        private List<SimpleArg> _SupportedSwitches;
        private List<SimpleArg> _OptionalArgsWithDefaultValue;
        public SimpleArgParser(List<SimpleArg> SupportedArgs)
        {
            _SupportedArgs = SupportedArgs;
            _RequiredArgs = _SupportedArgs.Where(sa => sa.Required == true).ToList();
            _OptionalArgsWithDefaultValue = _SupportedArgs.Where(sa => (sa.Required == false) && (sa.DefaultValue!=null)).ToList();
            _SupportedSwitches = _SupportedArgs.Where(sa => sa.IsSwitch == true).ToList();
        }

        public Dictionary<string, dynamic> ParseArgs(string[] Args)
        {
            List<string> SuppliedSwitches = new List<string>();
            List<string> argsList = Args.ToList<string>();

            //Remove Switches before key pair matching
            foreach (var supportedSwitch in _SupportedSwitches)
            {
                var ind = argsList.IndexOf(supportedSwitch.Name);
                if (ind >= 0)
                {
                    SuppliedSwitches.Add(supportedSwitch.Name);
                    argsList.RemoveAt(ind);
                }
            }

            //Validate argument formats
            Dictionary<string, dynamic> parsedArgs = new Dictionary<string, dynamic>();
            for (int i = 0; i < argsList.Count; i+=2)
            {
                if (parsedArgs.ContainsKey(argsList[i]))
                {
                    continue;
                }

                if(i<argsList.Count-1)
                {
                    var supportedArg = _SupportedArgs.Where(sa => sa.Name == argsList[i]).FirstOrDefault();
                    if(supportedArg==null)
                    {
                        throw new ArgumentException(String.Format("Argument not supported: {0}", argsList[i]));
                    }

                    CultureInfo provider = CultureInfo.InvariantCulture;

                    switch (supportedArg.ArgType)
                    {
                        case SimpleArgType.String:
                            {
                                parsedArgs.Add(argsList[i], argsList[i + 1]);
                                break;
                            }
                        case SimpleArgType.Filename:
                            {
                                if(!System.IO.Path.IsPathRooted(argsList[i + 1]))
                                {
                                    argsList[i + 1] = System.IO.Path.GetFullPath(argsList[i + 1]);
                                }
                               
                                if(!System.IO.File.Exists(argsList[i + 1]))
                                {
                                    throw new ArgumentException(String.Format("File not found: {0}", argsList[i + 1]));
                                }

                                parsedArgs.Add(argsList[i], argsList[i + 1]);
                                break;
                            }
                        case SimpleArgType.Integer:
                            {
                                int testI = 0;
                                if(!int.TryParse(argsList[i + 1], out testI))
                                {
                                    throw new ArgumentException(String.Format("Argument malformed. Expected integer: {0} -> {1}", argsList[i], argsList[i + 1]));
                                }
                                parsedArgs.Add(argsList[i], testI);
                                break;
                            }
                        case SimpleArgType.Decimal:
                            {
                                double testD = 0.0;
                                if (!double.TryParse(argsList[i + 1], NumberStyles.Float, provider, out testD))
                                {
                                    throw new ArgumentException(String.Format("Argument malformed. Expected decimal number: {0} -> {1}", argsList[i], argsList[i + 1]));
                                }
                                parsedArgs.Add(argsList[i], testD);
                                break;
                            }
                        case SimpleArgType.DateTime:
                            {
                                
                                DateTime dt = DateTime.MinValue;
                                string[] supportedDateTimeFormats = { "yyyy-MM-dd hh:mm:ss" };
                                bool success = false;
                                foreach (string s in supportedDateTimeFormats)
                                {
                                    if(DateTime.TryParseExact(argsList[i + 1], s, provider, DateTimeStyles.None, out dt))
                                    {
                                        success = true;
                                        break;
                                    }
                                }
                                    
                                if(!success)
                                {
                                    throw new ArgumentException(String.Format("Argument malformed. Expected DateTime: {0} in supported format. Supported formats {1}", argsList[i], String.Join(" / ", supportedDateTimeFormats)));
                                }
                      
                                parsedArgs.Add(argsList[i], dt);
                                break;
                            }
                        case SimpleArgType.Date:
                            {
                                DateTime dt = DateTime.MinValue;
                                string[] supportedDateTimeFormats = { "yyyy-MM-dd" };
                                bool success = false;
                                foreach (string s in supportedDateTimeFormats)
                                {
                                    if (DateTime.TryParseExact(argsList[i + 1], s, provider, DateTimeStyles.None, out dt))
                                    {
                                        success = true;
                                        break;
                                    }
                                }

                                if (!success)
                                {
                                    throw new ArgumentException(String.Format("Argument malformed. Expected Date: {0} in supported format. Supported formats {1}", argsList[i], String.Join(" / ", supportedDateTimeFormats)));
                                }

                                parsedArgs.Add(argsList[i], dt);
                                break;
                            }
                        case SimpleArgType.Time:
                            {
                                DateTime dt = DateTime.MinValue;
                                string[] supportedDateTimeFormats = { "HH:mm:ss" };
                                bool success = false;
                                foreach (string s in supportedDateTimeFormats)
                                {
                                    if (DateTime.TryParseExact(argsList[i + 1], s, provider, DateTimeStyles.None, out dt))
                                    {
                                        success = true;
                                        break;
                                    }
                                }

                                if (!success)
                                {
                                    throw new ArgumentException(String.Format("Argument malformed. Expected Time: {0} in supported format. Supported formats {1}", argsList[i], String.Join(" / ", supportedDateTimeFormats)));
                                }

                                parsedArgs.Add(argsList[i], dt);
                                break;
                            }
                        case SimpleArgType.Boolean:
                            {
                                string[] trues = { "true", "1" };
                                string[] falses = { "false", "0" };
                                bool testB = false;
                                if(Array.IndexOf<string>(trues, argsList[i])>=0)
                                {
                                    testB = true;
                                }
                                else if (Array.IndexOf<string>(falses, argsList[i]) >= 0)
                                {
                                    testB = false;
                                }
                                else
                                {
                                    throw new ArgumentException(String.Format("Argument malformed. Expected boolean in supported format: {0} -> true/1/false/0", argsList[i], argsList[i + 1]));
                                }

                                parsedArgs.Add(argsList[i], testB);
                                break;
                            }
                    }
                }
                else
                {
                    parsedArgs.Add(argsList[i], null);
                }
            }

            //Add the switches back to the end
            foreach(string s in SuppliedSwitches)
            {
                parsedArgs.Add(s, null);
            }

            //Add defaults for optional args not passed
            foreach(var optionalArg in _OptionalArgsWithDefaultValue)
            {
                if(!parsedArgs.ContainsKey(optionalArg.Name))
                {
                    parsedArgs.Add(optionalArg.Name, optionalArg.DefaultValue);
                }
            }


            //Check that required arguments were passed
            if (_RequiredArgs != null)
            {
                foreach (var pa in _RequiredArgs)
                {
                    if (!parsedArgs.ContainsKey(pa.Name))
                    {
                        throw new ArgumentException(String.Format("Required argument {0} not provided", pa.Name));
                    }
                }
            }

            //Check that no unsupported arguments passed
            foreach (var arg in parsedArgs)
            {
                var sa = this._SupportedArgs.Where(a => a.Name == arg.Key).FirstOrDefault();
                if (sa == null)
                {
                    throw new ArgumentException(String.Format("Argument not recognised: {0}", arg.Key));
                }
            }

            return parsedArgs;
        }

        public void ShowUsage()
        {
            //_SupportedArgs

            Console.WriteLine(String.Empty);
            StringBuilder sb = new StringBuilder();
            sb.Append(String.Format("Usage: {0}", System.AppDomain.CurrentDomain.FriendlyName));

            foreach( var arg in this._SupportedArgs)
            {
                if (arg.Required)
                {
                    sb.Append(String.Format(" {0} {1}", arg.Name, arg.ExmaplePlaceholder));
                }
                else
                {
                    sb.Append(String.Format(" {0} [{1}]", arg.Name, arg.ExmaplePlaceholder));
                }
            }

            Console.WriteLine(sb.ToString());
            Console.WriteLine(Environment.NewLine);

            foreach (var arg in this._SupportedArgs)
            {
                Console.WriteLine(String.Format("{0}:\t\t{1}", arg.Name, arg.Description));
            }

                //Console.WriteLine(String.Format("Usage: {0} -filename [excelfilename] -worksheet [1-n] -format [html|wikitable|jsonarrays|jsonobjects]", System.AppDomain.CurrentDomain.FriendlyName));


            //    Console.WriteLine("-filename: The Excel file name, relative or absolute path.");
            //Console.WriteLine("-worksheet is optional. Defaults to 1");
            //Console.WriteLine("-format [optional]. Output file format. Defaults to wikitable");
            //Console.WriteLine("\t\tjsonarrays option outputs each row as an array, making a 2-d array.");
            //Console.WriteLine("\t\tjsonobjects option outputs an array of objects using the first row as object key names");
            //Console.WriteLine("-range [optional]. Excel worksheet range to export. Defaults to used range");
            Console.WriteLine(Environment.NewLine);
        }
    }

    public enum SimpleArgType { String = 0 , Filename, Integer, Decimal, Date, DateTime, Time, Boolean }

    public class SimpleArg
    {
        public bool IsSwitch { get; set; }

        public string Name { get; set; }

        public bool Required { get; set; }

        public dynamic DefaultValue { get; set; }

        public string Description { get; set; }

        public string ExmaplePlaceholder { get; set; }

        public SimpleArgType ArgType { get; set; }
    }

    public class PassedArg
    {
        public string Name { get; }
        public dynamic Value { get; }
    }
}
