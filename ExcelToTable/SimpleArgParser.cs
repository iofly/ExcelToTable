using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;

namespace ExcelToTable
{
	public class SimpleArgParser
	{
		private readonly List<SimpleArg> _supportedArgs;
		private readonly List<SimpleArg> _requiredArgs;
		private readonly List<SimpleArg> _supportedSwitches;
		private readonly List<SimpleArg> _optionalArgsWithDefaultValue;

		public Dictionary<string, dynamic> ParsedArguments { get; set; }

		public SimpleArgParser(List<SimpleArg> supportedArgs, string[] args)
		{
			_supportedArgs = supportedArgs;
			_requiredArgs = _supportedArgs.Where(sa => sa.Required).ToList();
			_optionalArgsWithDefaultValue = _supportedArgs.Where(sa => (sa.Required == false) && (sa.DefaultValue!=null)).ToList();
			_supportedSwitches = _supportedArgs.Where(sa => sa.IsSwitch).ToList();
			ParsedArguments = ParseArgs(args);
		}

		private Dictionary<string, dynamic> ParseArgs(string[] args)
		{
			var suppliedSwitches = new List<string>();
			var argsList = args.ToList();

			//Remove Switches before key pair matching
			foreach (var supportedSwitch in _supportedSwitches)
			{
				var ind = argsList.IndexOf(supportedSwitch.Name);
				if (ind >= 0)
				{
					suppliedSwitches.Add(supportedSwitch.Name);
					argsList.RemoveAt(ind);
				}
			}

			//Validate argument formats
			Dictionary<string, dynamic> parsedArgs = new Dictionary<string, dynamic>();
			for (int i = 0; i < argsList.Count; i+=2)
			{
				if (parsedArgs.ContainsKey(argsList[i]))
				{
					throw new ArgumentException($"Duplicate argument {argsList[i]}");
				}

				if(i<argsList.Count-1)
				{
					var supportedArg = _supportedArgs.FirstOrDefault(sa => sa.Name == argsList[i]);
					if(supportedArg==null)
					{
						throw new ArgumentException($"Argument not supported: {argsList[i]}");
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
								if (!Path.IsPathRooted(argsList[i + 1]))
								{
									argsList[i + 1] = Path.GetFullPath(argsList[i + 1]);
								}

								if (!IsValidFilename(argsList[i + 1]))
								{
									throw new ArgumentException($"Invalid filename for argument: {argsList[i]} => '{argsList[i + 1]}'");
								}

								parsedArgs.Add(argsList[i], argsList[i + 1]);

								break;
							}
						case SimpleArgType.NewFilename:
							{
								if(!Path.IsPathRooted(argsList[i + 1]))
								{
									argsList[i + 1] = Path.GetFullPath(argsList[i + 1]);
								}

								if(!IsValidFilename(argsList[i + 1]))
								{
									throw new ArgumentException($"Invalid filename for argument: {argsList[i]} => '{argsList[i + 1]}'");
								}
								else if (File.Exists(argsList[i + 1]))
								{
									throw new ArgumentException($"File already exists: {argsList[i]} => '{argsList[i + 1]}'");
								}

								parsedArgs.Add(argsList[i], argsList[i + 1]);

								break;
							}
						case SimpleArgType.ExistingFilename:
							{
								if (!Path.IsPathRooted(argsList[i + 1]))
								{
									argsList[i + 1] = Path.GetFullPath(argsList[i + 1]);
								}

								if (!File.Exists(argsList[i + 1]))
								{
									throw new ArgumentException($"File not found: {argsList[i]} => '{argsList[i + 1]}'");
								}

								parsedArgs.Add(argsList[i], argsList[i + 1]);

								break;
							}
						case SimpleArgType.Integer:
							{
								if(!int.TryParse(argsList[i + 1], out var testI))
								{
									throw new ArgumentException($"Argument malformed. Expected integer: {argsList[i]} => '{argsList[i + 1]}'");
								}

								parsedArgs.Add(argsList[i], testI);

								break;
							}
						case SimpleArgType.Decimal:
							{
								if (!double.TryParse(argsList[i + 1], NumberStyles.Float, provider, out var testD))
								{
									throw new ArgumentException(
										$"Argument malformed. Expected decimal number: {argsList[i]} => '{argsList[i + 1]}'");
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
									throw new ArgumentException(
										$"DateTime argument malformed. Argument {argsList[i]} => '{argsList[i + 1]}' is not in a supported format. Supported formats: [{String.Join(", ", supportedDateTimeFormats)}]");
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
									throw new ArgumentException(
										$"Date argument malformed. Argument {argsList[i]} => '{argsList[i + 1]}' is not in a supported format. Supported formats: [{String.Join(", ", supportedDateTimeFormats)}]");
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
									throw new ArgumentException(
										$"Time argument malformed. Argument {argsList[i]} => '{argsList[i + 1]}' is not in a supported format. Supported formats: [{String.Join(", ", supportedDateTimeFormats)}]");
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
									throw new ArgumentException(
										$"Boolean argument malformed. Argument {argsList[i]} => '{argsList[i + 1]}' is not in a supported format. Supported formats: [{String.Join(", ", trues)}, {String.Join(", ", falses)}]");
								}
								
								break;
							}
						case SimpleArgType.Uri:
							{
								if(Uri.TryCreate(argsList[i + 1], UriKind.Absolute, out var uri))
								{
									parsedArgs.Add(argsList[i], uri);
								}
								else
								{
									throw new ArgumentException(
										$"URI argument malformed. Argument {argsList[i]} => '{argsList[i + 1]}' is not in a supported URI format.");
								}
								break;
							}
						case SimpleArgType.EmailAddress:
							{
								if (Regex.IsMatch(argsList[i + 1], @"^([\w\.\-]+)@([\w\-]+)((\.(\w){ 2,3})+)$"))
								{
									parsedArgs.Add(argsList[i], argsList[i + 1]);
								}
								else
								{
									throw new ArgumentException(
										$"Email argument malformed. Argument {argsList[i]} => '{argsList[i + 1]}' is not a valid email address.");
								}

								break;
							}
						case SimpleArgType.Guid:
							{
								if(Regex.IsMatch(argsList[i + 1], @"^[{(]?[0-9A-F]{8}[-]?([0-9A-F]{4}[-]?){3}[0-9A-F]{12}[)}]?$", RegexOptions.IgnoreCase))
								{ 
									parsedArgs.Add(argsList[i], argsList[i + 1]);
								}
								else
								{
									throw new ArgumentException(
										$"GUID argument malformed. Argument {argsList[i]} => '{argsList[i + 1]}' is not a valid GUID.");
								}

								break;
							}
						case SimpleArgType.ExcelRange:
							{
								var wsrcs = ParseExcelRange(argsList[i + 1]);
								if (wsrcs == null)
								{
									throw new ArgumentException($"Worksheet range parameter is not valid: '{argsList[i + 1]}'");
								}
								else
								{
									parsedArgs.Add(argsList[i], wsrcs);
								}

								break;
							}
						case SimpleArgType.ValueRange:
							{
								var sa = _supportedArgs.FirstOrDefault(ar => ar.Name == argsList[i]);
								if (sa == null)
								{
									throw new ArgumentException($"ValueRange parameter is not valid: {argsList[i]} = {argsList[i + 1]}");
								}
								else if (sa.ValueRange.IndexOf(argsList[i + 1]) < 0)
								{
									throw new ArgumentException(
										$"Value supplied for {argsList[i]} is not in the valid range of values [{String.Join(", ", sa.ValueRange)}]");
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
			foreach(string s in suppliedSwitches)
			{
				parsedArgs.Add(s, null);
			}

			//Check the inc/exc list:
			//If an arg is optional, its exclusion/inclusion list is ignored
			//If an arg is required it will have no effect if included in an exclusion or inclusion list
			var argumentsWithExclusionList = _requiredArgs.Where(a => (a.ExcludeArgs != null) && (a.Required)).ToList();
			var argumentsWithInclusionList = _requiredArgs.Where(a => (a.IncludeArgs != null) && (a.Required)).ToList();
		   
			//Check exclusion
			foreach (var arg in parsedArgs)
			{
				var argExcludeCheck = argumentsWithExclusionList.FirstOrDefault(fa => fa.ExcludeArgs.Contains(arg.Key));
				if(argExcludeCheck!=null)
				{
					throw new ArgumentException(
						$"Argument '{arg.Key}' cannot be passed if argument '{argExcludeCheck.Name}' has been passed.");
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
						throw new ArgumentException($"Argument '{s}' must be passed if argument '{argI.Name}' has been passed.");
					}
				}
			}

			//Check that required arguments were passed
			if (_requiredArgs != null)
			{
				foreach (var pa in _requiredArgs)
				{
					if (!parsedArgs.ContainsKey(pa.Name))
					{
						throw new ArgumentException($"Required argument {pa.Name} not provided");
					}
				}
			}

			//Check that no unsupported arguments passed
			foreach (var arg in parsedArgs)
			{
				var sa = _supportedArgs.FirstOrDefault(a => a.Name == arg.Key);
				if (sa == null)
				{
					throw new ArgumentException($"Argument not recognised: {arg.Key}");
				}
			}

			//Add defaults for optional args not passed
			foreach (var optionalArg in _optionalArgsWithDefaultValue)
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
			ShowUsage(_supportedArgs);
		}

		public static void ShowUsage(List<SimpleArg> supportedArgs)
		{
			Console.WriteLine(String.Empty);
			StringBuilder sb = new StringBuilder();
			sb.Append($"Usage: {AppDomain.CurrentDomain.FriendlyName}");

			foreach (var arg in supportedArgs)
			{
				sb.Append(arg.Required ? $" {arg.Name} {arg.ExmaplePlaceholder}" : $" {arg.Name} [{arg.ExmaplePlaceholder}]");
			}

			Console.WriteLine(sb.ToString());
			Console.WriteLine(Environment.NewLine);

			foreach (var arg in supportedArgs)
			{
				Console.WriteLine($"{arg.Name}:\t\t{arg.Description}");
			}

			Console.WriteLine(Environment.NewLine);
		}

		private bool IsValidFilename(string filename)
		{
			if (((!string.IsNullOrEmpty(filename)) && (filename.IndexOfAny(Path.GetInvalidPathChars()) >= 0)) == false)
			{
				return false;
			}

			//check for 2 seperators in a row in filename, except the start
			string sep2 = String.Format("{0}{0}", Path.DirectorySeparatorChar);
			if (filename.IndexOf(sep2, StringComparison.Ordinal)>=1) //check fron 2nd char on, don't disallow \\ at start of filename, its a valid UNC path root
			{
				return false;
			}

			//Check for trailing spaces in dir path elements
			char[] sep = { Path.DirectorySeparatorChar };
			string[] parts = filename.Split(sep);

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

		public static WorkSheetRangeCoordinates ParseExcelRange(string range)
		{
			const string letters = "_ABCDEFGHIJKLMNOPQRSTUVWXYZ";
			//BMZ4:BNC14
			string regexPattern = @"([A-Z]{1,3})(\d{1,7}):([A-Z]{1,3})(\d{1,7})";
			WorkSheetRangeCoordinates rangeCoords = new WorkSheetRangeCoordinates();
			var matches = Regex.Matches(range, regexPattern);
			int colIndex = 0;

			if (matches.Count > 0)
			{
				int letterIndex;
				string part1 = matches[0].Groups[1].Value;
				char[] part1X = part1.ToCharArray();
				Array.Reverse(part1X); //make first car the LSC
				for (int i = 0; i < part1X.Length; i++)
				{
					letterIndex = letters.IndexOf(part1X[i]);
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
					letterIndex = letters.IndexOf(part2X[i]);
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
		Uri,
		EmailAddress,
		Guid,
		ExcelRange,
		ValueRange, 
        Switch
	}

	public class SimpleArg
	{
		public SimpleArg()
		{
			ExcludeArgs = new List<string>();
			IncludeArgs = new List<string>();
			ValueRange = new List<string>();
		}

		public bool IsSwitch { get; set; }

		public string Name { get; set; }

		public bool Required { get; set; }

		public dynamic DefaultValue { get; set; }

		public string Description { get; set; }

		public string ExmaplePlaceholder { get; set; }

		public SimpleArgType ArgType { get; set; }

		#region Public Accessors

		/// <summary>
		/// List of argument names that MUST NOT be passed if this argument is passed.
		/// </summary>
		public List<string> ExcludeArgs { get; set; }

		/// <summary>
		/// List of argument names that MUST be passed if this argument is passed.
		/// </summary>
		public List<string> IncludeArgs { get; set; }

		/// <summary>
		/// List of values that the parameter is permitted to have
		/// </summary>
		public List<string> ValueRange { get; set; }

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
			TopLeft = new WorkSheetCoordinate();
			BottomRight = new WorkSheetCoordinate();
		}

		public WorkSheetCoordinate TopLeft { get; set; }

		public WorkSheetCoordinate BottomRight { get; set; }
	}
}
