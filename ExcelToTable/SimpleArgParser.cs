using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Text.RegularExpressions;

namespace SimpleArgs
{
	public class SimpleArgParser
	{
		private List<SimpleArg> _SupportedArgs;
		private List<SimpleArg> _RequiredArgs;
		private List<SimpleArg> _SupportedSwitches;
		private List<SimpleArg> _OptionalArgsWithDefaultValue;
		private Dictionary<string, dynamic> _ParsedArguments;

		public Dictionary<string, dynamic> ParsedArguments
		{
			get
			{
				return _ParsedArguments;
			}
			set
			{
				_ParsedArguments = value;
			}
		}

		public SimpleArgParser(List<SimpleArg> SupportedArgs, string[] args)
		{
			_SupportedArgs = SupportedArgs;
			_RequiredArgs = _SupportedArgs.Where(sa => sa.Required == true).ToList();
			_OptionalArgsWithDefaultValue = _SupportedArgs.Where(sa => (sa.Required == false) && (sa.DefaultValue!=null)).ToList();
			_SupportedSwitches = _SupportedArgs.Where(sa => sa.IsSwitch == true).ToList();
			_ParsedArguments = this.ParseArgs(args);
		}

		private Dictionary<string, dynamic> ParseArgs(string[] Args)
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
					throw new ArgumentException(String.Format("Duplicate argument {0}", argsList[i]));
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
						case SimpleArgType.FileName:
							{
								if (!System.IO.Path.IsPathRooted(argsList[i + 1]))
								{
									argsList[i + 1] = System.IO.Path.GetFullPath(argsList[i + 1]);
								}

								if (!IsValidFilename(argsList[i + 1]))
								{
									throw new ArgumentException(String.Format("Invalid filename for argument: {0} => '{1}'", argsList[i], argsList[i + 1]));
								}

								parsedArgs.Add(argsList[i], argsList[i + 1]);

								break;
							}
						case SimpleArgType.NewFilename:
							{
								if(!System.IO.Path.IsPathRooted(argsList[i + 1]))
								{
									argsList[i + 1] = System.IO.Path.GetFullPath(argsList[i + 1]);
								}

								if(!IsValidFilename(argsList[i + 1]))
								{
									throw new ArgumentException(String.Format("Invalid filename for argument: {0} => '{1}'", argsList[i], argsList[i + 1]));
								}
								else if (System.IO.File.Exists(argsList[i + 1]))
								{
									throw new ArgumentException(String.Format("File already exists: {0} => '{1}'", argsList[i], argsList[i + 1]));
								}

								parsedArgs.Add(argsList[i], argsList[i + 1]);

								break;
							}
						case SimpleArgType.ExistingFilename:
							{
								if (!System.IO.Path.IsPathRooted(argsList[i + 1]))
								{
									argsList[i + 1] = System.IO.Path.GetFullPath(argsList[i + 1]);
								}

								if (!System.IO.File.Exists(argsList[i + 1]))
								{
									throw new ArgumentException(String.Format("File not found: {0} => '{1}'", argsList[i], argsList[i + 1]));
								}

								parsedArgs.Add(argsList[i], argsList[i + 1]);

								break;
							}
						case SimpleArgType.Integer:
							{
								if(!int.TryParse(argsList[i + 1], out var testI))
								{
									throw new ArgumentException(String.Format("Argument malformed. Expected integer: {0} => '{1}'", argsList[i], argsList[i + 1]));
								}

								parsedArgs.Add(argsList[i], testI);

								break;
							}
						case SimpleArgType.Decimal:
							{
								if (!double.TryParse(argsList[i + 1], NumberStyles.Float, provider, out var testD))
								{
									throw new ArgumentException(String.Format("Argument malformed. Expected decimal number: {0} => '{1}'", argsList[i], argsList[i + 1]));
								}

								parsedArgs.Add(argsList[i], testD);

								break;
							}
						case SimpleArgType.DateTime:
							{
								
								DateTime dt = DateTime.MinValue;
								string[] supportedDateTimeFormats = { "yyyy-MM-dd hh:mm:ss", "yyyy/MM/dd hh:mm:ss" };
								bool success = false;
								foreach (string s in supportedDateTimeFormats)
								{
									if(DateTime.TryParseExact(argsList[i + 1], s, provider, DateTimeStyles.None, out dt))
									{
										success = true;
										break;
									}
								}

								if (!success)
								{
									throw new ArgumentException(String.Format("DateTime argument malformed. Argument {0} => '{1}' is not in a supported format. Supported formats: [{2}]", argsList[i], argsList[i + 1], String.Join(", ", supportedDateTimeFormats)));
								}
								else
								{
									parsedArgs.Add(argsList[i], dt);
								}
								break;
							}
						case SimpleArgType.Date:
							{
								DateTime dt = DateTime.MinValue;
								string[] supportedDateTimeFormats = { "yyyy-MM-dd", "yyyy/MM/dd" };
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
									throw new ArgumentException(String.Format("Date argument malformed. Argument {0} => '{1}' is not in a supported format. Supported formats: [{2}]", argsList[i], argsList[i + 1], String.Join(", ", supportedDateTimeFormats)));
								}
								else
								{
									parsedArgs.Add(argsList[i], dt);
								}

								break;
							}
						case SimpleArgType.Time:
							{
								DateTime dt = DateTime.MinValue;
								string[] supportedDateTimeFormats = { "HH:mm:ss", "HH:mm" };
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
									throw new ArgumentException(String.Format("Time argument malformed. Argument {0} => '{1}' is not in a supported format. Supported formats: [{2}]", argsList[i], argsList[i + 1], String.Join(", ", supportedDateTimeFormats)));
								}
								else
								{
									parsedArgs.Add(argsList[i], dt);
								}

								break;
							}
						case SimpleArgType.Boolean:
							{
								string[] trues = { "true", "1", "yes", "y" };
								string[] falses = { "false", "0", "no", "n" };

								if (trues.Contains(argsList[i + 1]))
								{
									parsedArgs.Add(argsList[i], true);
								}
								else if (falses.Contains(argsList[i + 1]))
								{
									parsedArgs.Add(argsList[i], false);
								}
								else
								{
									throw new ArgumentException(String.Format("Boolean argument malformed. Argument {0} => '{1}' is not in a supported format. Supported formats: [{2}, {3}]", argsList[i], argsList[i + 1], String.Join(", ", trues), String.Join(", ", falses)));
								}
								
								break;
							}
						case SimpleArgType.URI:
							{
								if(Uri.TryCreate(argsList[i + 1], UriKind.Absolute, out var uri))
								{
									parsedArgs.Add(argsList[i], uri);
								}
								else
								{
									throw new ArgumentException(String.Format("URI argument malformed. Argument {0} => '{1}' is not in a supported URI format.", argsList[i], argsList[i + 1]));
								}
								break;
							}
						case SimpleArgType.EmailAddress:
							{
								if (System.Text.RegularExpressions.Regex.IsMatch(argsList[i + 1], @"^([\w\.\-]+)@([\w\-]+)((\.(\w){ 2,3})+)$"))
								{
									parsedArgs.Add(argsList[i], argsList[i + 1]);
								}
								else
								{
									throw new ArgumentException(String.Format("Email argument malformed. Argument {0} => '{1}' is not a valid email address.", argsList[i], argsList[i + 1]));
								}

								break;
							}
						case SimpleArgType.Guid:
							{
								if(Guid.TryParse(argsList[i + 1], out var g))
								{
									parsedArgs.Add(argsList[i], argsList[i + 1]);
								}
								else
								{
									throw new ArgumentException(String.Format("GUID argument malformed. Argument {0} => '{1}' is not a valid GUID.", argsList[i], argsList[i + 1]));
								}

								break;
							}
						case SimpleArgType.ExcelRange:
							{
								var wsrcs = ParseExcelRange(argsList[i + 1]);
								if (wsrcs == null)
								{
									throw new ArgumentException(String.Format("Worksheet range parameter is not valid: '{0}'", argsList[i + 1]));
								}
								else
								{
									parsedArgs.Add(argsList[i], wsrcs);
								}

								break;
							}
						case SimpleArgType.ValueRange:
							{
								var sa = _SupportedArgs.Where(ar => ar.Name == argsList[i]).FirstOrDefault();
								if (sa == null)
								{
									throw new ArgumentException(String.Format("ValueRange parameter is not valid: {0} = {1}", argsList[i], argsList[i + 1]));
								}
								else if (sa.ValueRange.IndexOf(argsList[i + 1]) < 0)
								{
									throw new ArgumentException(String.Format("Value supplied for {0} is not in the valid range of values [{1}]", argsList[i], String.Join(", ", sa.ValueRange)));
								}
								else
								{
									parsedArgs.Add(argsList[i], argsList[i + 1]);
								}

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

			//Check the inc/exc list:
			//If an arg is optional, its exclusion/inclusion list is ignored
			//If an arg is required it will have no effect if included in an exclusion or inclusion list
			var argumentsWithExclusionList = this._RequiredArgs.Where(a => (a.ExcludeArgs != null) && (a.Required == true)).ToList();
			var argumentsWithInclusionList = this._RequiredArgs.Where(a => (a.IncludeArgs != null) && (a.Required == true)).ToList();
		   
			//Check exclusion
			foreach (var arg in parsedArgs)
			{
				var argExcludeCheck = argumentsWithExclusionList.Where(fa => fa.ExcludeArgs.Contains(arg.Key)).FirstOrDefault();
				if(argExcludeCheck!=null)
				{
					throw new ArgumentException(String.Format("Argument '{0}' cannot be passed if argument '{1}' has been passed.", arg.Key, argExcludeCheck.Name));
				}
			}
		   
			//Check Inclusion
			foreach (var argI in argumentsWithInclusionList)
			{
				if(!parsedArgs.ContainsKey(argI.Name))
				{
					continue;
				}

				//was every arg in argI.IncludeArgs privided?
				foreach(string s in argI.IncludeArgs)
				{
					if (!parsedArgs.ContainsKey(s))
					{
						throw new ArgumentException(String.Format("Argument '{0}' must be passed if argument '{1}' has been passed.", s, argI.Name));
					}
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

			//Add defaults for optional args not passed
			foreach (var optionalArg in _OptionalArgsWithDefaultValue)
			{
				if(!parsedArgs.ContainsKey(optionalArg.Name))
				{
					parsedArgs.Add(optionalArg.Name, optionalArg.DefaultValue);
				}
			}

			return parsedArgs;
		}

		public void ShowUsage()
		{
			SimpleArgParser.ShowUsage(this._SupportedArgs);
		}

		public static void ShowUsage(List<SimpleArg> SupportedArgs)
		{
			Console.WriteLine(String.Empty);
			StringBuilder sb = new StringBuilder();
			sb.Append(String.Format("Usage: {0}", System.AppDomain.CurrentDomain.FriendlyName));

			foreach (var arg in SupportedArgs)
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

			foreach (var arg in SupportedArgs)
			{
				Console.WriteLine(String.Format("{0}:\t\t{1}", arg.Name, arg.Description));
			}

			Console.WriteLine(Environment.NewLine);
		}

		private bool IsValidFilename(string Filename)
		{
			if (((!string.IsNullOrEmpty(Filename)) && (Filename.IndexOfAny(System.IO.Path.GetInvalidPathChars()) >= 0)) == false)
			{
				return false;
			}

			//check for 2 seperators in a row in filename, except the start
			string sep2 = String.Format("{0}{0}", System.IO.Path.DirectorySeparatorChar);
			if (Filename.IndexOf(sep2)>=1) //check fron 2nd char on, don't disallow \\ at start of filename, its a valid UNC path root
			{
				return false;
			}

			//Check for trailing spaces in dir path elements
			char[] sep = { System.IO.Path.DirectorySeparatorChar };
			string[] parts = Filename.Split(sep);

			foreach(string s in parts)
			{
				if(s.Trim()!=s)
				{
					return false;
				}
				if(String.IsNullOrWhiteSpace(s))
				{
					return false;
				}
			}

			return true;
		}

		private bool DirPartExists(string Filename, out string AbsDirName)
		{
			if(System.IO.Path.IsPathRooted(Filename))
			{
				Filename = System.IO.Path.GetFullPath(Filename);
			}

			AbsDirName = System.IO.Path.GetDirectoryName(Filename);

			return System.IO.Directory.Exists(AbsDirName);
		}

		public static WorkSheetRangeCoordinates ParseExcelRange(string Range)
		{
			const string Letters = "_ABCDEFGHIJKLMNOPQRSTUVWXYZ";
			//BMZ4:BNC14
			string regexPattern = @"([A-Z]{1,3})(\d{1,7}):([A-Z]{1,3})(\d{1,7})";
			WorkSheetRangeCoordinates rangeCoords = new WorkSheetRangeCoordinates();
			var matches = Regex.Matches(Range, regexPattern);

			int colIndex = 0;
			int letterIndex = 0;

			if (matches.Count > 0)
			{

				string part1 = matches[0].Groups[1].Value;
				char[] part1X = part1.ToCharArray();
				Array.Reverse(part1X); //make first car the LSC
				for (int i = 0; i < part1X.Length; i++)
				{
					letterIndex = Letters.IndexOf(part1X[i]);
					colIndex += letterIndex * Convert.ToInt32(Math.Pow(26, i));
				}
				rangeCoords.TopLeft.X = colIndex;
				rangeCoords.TopLeft.Y = int.Parse(matches[0].Groups[2].Value);

				colIndex = 0;

				string part2 = matches[0].Groups[3].Value;
				char[] part2X = part2.ToCharArray();
				Array.Reverse(part2X); //make first car the LSC
				for (int i = 0; i < part2X.Length; i++)
				{
					letterIndex = Letters.IndexOf(part2X[i]);
					colIndex += letterIndex * Convert.ToInt32(Math.Pow(26, i));
				}
				rangeCoords.BottomRight.X = colIndex;
				rangeCoords.BottomRight.Y = int.Parse(matches[0].Groups[4].Value);

				if (!RangeIsValid(rangeCoords))
				{
					return null;
				}


				return rangeCoords;
			}
			else
			{
				return null;
			}
		}

		public static bool RangeIsValid(WorkSheetRangeCoordinates wsrc)
		{
			//Excel limits = 1,048,576 rows by 16,384 columns
			if (wsrc == null)
			{
				return false;
			}

			return (wsrc.TopLeft.X <= 16384) && (wsrc.BottomRight.X <= 16384) && (wsrc.TopLeft.Y <= 1048576) && (wsrc.BottomRight.Y <= 1048576);

			//long l = Math.Abs(Convert.ToInt64(wsrc.TopLeft.X - wsrc.BottomRight.X)) * Math.Abs(Convert.ToInt64(wsrc.TopLeft.Y - wsrc.BottomRight.Y));
		}
	}

	public enum SimpleArgType
    {
        String = 0,
        FileName,
        NewFilename,
        ExistingFilename,
        Integer,
        Decimal,
        Date,
        DateTime,
        Time,
        Boolean,
        URI,
        EmailAddress,
        Guid,
        ExcelRange,
        ValueRange
    }

	public class SimpleArg
	{
		public SimpleArg()
		{
			_ExcludeArgs = new List<string>();
			_IncludeArgs = new List<string>();
			_ValueRange = new List<string>();
		}

        public bool IsSwitch { get; set; }

		public string Name { get; set; }

		public bool Required { get; set; }

		public dynamic DefaultValue { get; set; }

		public string Description { get; set; }

		public string ExmaplePlaceholder { get; set; }

		public SimpleArgType ArgType { get; set; }

        #region Public Accessors
        private List<string> _ExcludeArgs;
		/// <summary>
		/// List of argument names that MUST NOT be passed if this argument is passed.
		/// </summary>
		public List<string> ExcludeArgs
		{
			get
			{
				return _ExcludeArgs;
			}
			set
			{
				_ExcludeArgs = value;
			}
		}

		private List<string> _IncludeArgs;
		/// <summary>
		/// List of argument names that MUST be passed if this argument is passed.
		/// </summary>
		public List<string> IncludeArgs
		{
			get
			{
				return _IncludeArgs;
			}
			set
			{
				_IncludeArgs = value;
			}
		}

		private List<string> _ValueRange;
		/// <summary>
		/// List of values that the parameter is permitted to have
		/// </summary>
		public List<string> ValueRange
		{
			get
			{
				return _ValueRange;
			}
			set
			{
				_ValueRange = value;
			}
		}
        #endregion
    }

	public class WorkSheetCoordinate
	{
		public int X { get; set; }
		public int Y { get; set; }
	}

	public class WorkSheetRangeCoordinates
	{
		public WorkSheetRangeCoordinates()
		{
			_TopLeft = new WorkSheetCoordinate();
			_BottomRight = new WorkSheetCoordinate();
		}

		private WorkSheetCoordinate _TopLeft;
		public WorkSheetCoordinate TopLeft
		{
			get
			{
				return _TopLeft;
			}
			set
			{
				_TopLeft = value;
			}
		}

		private WorkSheetCoordinate _BottomRight;
		public WorkSheetCoordinate BottomRight
		{
			get
			{
				return _BottomRight;
			}
			set
			{
				_BottomRight = value;
			}
		}
	}
}
