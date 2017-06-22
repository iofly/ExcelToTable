using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;

namespace ExcelToTable
{
    public static class SimpleArgParser
    {
        public static Dictionary<string, dynamic> ParseArgs(string[] Args, List<SimpleArg> SupportedArgs)
        {
            List<string> SuppliedSwitches = new List<string>();
            List<SimpleArg> supportedSwitches = SupportedArgs.Where(sa => sa.IsSwitch == true).ToList();
            List<string> argsList = Args.ToList<string>();

            foreach (var supportedSwitch in supportedSwitches)
            {
                var ind = argsList.IndexOf(supportedSwitch.Name);
                if (ind >= 0)
                {
                    SuppliedSwitches.Add(supportedSwitch.Name);
                    argsList.RemoveAt(ind);
                }
            }

            Dictionary<string, dynamic> parsedArgs = new Dictionary<string, dynamic>();

            for (int i = 0; i < argsList.Count; i+=2)
            {
                if(parsedArgs.ContainsKey(argsList[i])) continue;

                if(i<argsList.Count-1)
                {
                    var supportedArg = SupportedArgs.Where(sa => sa.Name == argsList[i]).FirstOrDefault();
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
                                    throw new ArgumentException(String.Format("Argument malformed. Expected DateTime: {0} in supported format ", argsList[i], String.Join(" / ", supportedDateTimeFormats)));
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
                                    throw new ArgumentException(String.Format("Argument malformed. Expected Date: {0} in supported format ", argsList[i], String.Join(" / ", supportedDateTimeFormats)));
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
                                    throw new ArgumentException(String.Format("Argument malformed. Expected Time: {0} in supported format ", argsList[i], String.Join(" / ", supportedDateTimeFormats)));
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


            return parsedArgs;
        }

        public static List<string> GetMatches(string s)
        {
            List<string> matchValues = new List<string>();
            MatchCollection matches = Regex.Matches(s, "\\\"(.|\\n)*?\\\"", RegexOptions.None);
            foreach (Match match in matches)
            {
                if (match.Success)
                {
                    matchValues.Add(match.Value);
                }
            }

            matches = Regex.Matches(s, "\\'(.|\\n)*?\\'", RegexOptions.None);
            foreach (Match match in matches)
            {
                if (match.Success)
                {
                    matchValues.Add(match.Value);
                }
            }
            return matchValues;
        }
    }

    public enum SimpleArgType { String = 0 , Integer, Decimal, Date, DateTime, Time, Boolean }

    public class SimpleArg
    {
        public bool IsSwitch { get; set; }

        public string Name { get; set; }

        public bool Required { get; set; }

        public dynamic DefaultValue { get; set; }

        public SimpleArgType ArgType { get; set; }
    }

    public class PassedArg
    {
        public string Name { get; }
        public dynamic Value { get; }
    }

}
