//METRONOME
//Written by NightWatchman
//Used to manipulate GPO Active Directory settings

//USAGE:

//To query for the existing properties of a GPO in Active Directory, use: metronome.exe query <Target GUID> 

//To add the necessary pieces to add a scheduled task to a GPO, use: metronome.exe schtasks <Target GUID> <LDAP Path to GPO> <Original gpcMachineExtensionNames> <Original versionnumber>

//To only increment the versionnumber field, use: metronome.exe versionset <Target GUID> <LDAP Path to GPO> <Original versionnumber>

//To revert a GPO to its original properties or set the properties to different values, use: metronome.exe set <Target GUID> <LDAP Path to GPO> <Original gpcMachineExtensionNames> <Original versionnumber>


using System;
using System.Text;
using System.Collections.Generic;
using System.Security.Principal;
using System.DirectoryServices;
using System.IO;
using System.Text.RegularExpressions;
using System.Linq;

public class GPOedit
{
	public static void Main(string[] args)
	{
		int argumentsize = args.Length;
		string help = "To query for the existing properties of a GPO in Active Directory, use: \n- metronome.exe query <Target GUID> \n\nTo add an immediate scheduled task to a GPO, use:\n- metronome.exe itask <Target GUID> <LDAP Path to GPO> <Original gpcMachineExtensionNames> <Original versionnumber> <task author> <task name> <command> <arguments>\n\nTo deploy a file using a GPO, use:\n- metronome.exe file <Target GUID> <LDAP Path to GPO> <Original gpcMachineExtensionNames> <Original versionnumber> <source file and path> <destination file> <destination path (with no trailing \\)>\n\nTo revert a GPO to its original properties or set the properties to different values, use:\n- metronome.exe set <Target GUID> <LDAP Path to GPO> <Original gpcMachineExtensionNames> <Original versionnumber>\n\nTo only increment the versionnumber field, use:\n- metronome.exe versionset <Target GUID> <LDAP Path to GPO> <Original versionnumber>\n\nTo enable or disable a GPO, use:\n- metronome.exe enable\\disable <Target GUID> <LDAP Path to GPO> <Original versionnumber>";
		string guid_filter = "";
		string guid_LDAP = "";
		string gpcMacExtName = "";
		string new_gpcMacExtName = "";
		string schtask_gpcMacExtName = "[{AADCED64-746C-4633-A97C-D61349046527}{CAB54552-DEEA-4691-817E-ED4A4D1AFC72}]";
		string file_gpcMacExtName = "[{7150F9BF-48AD-4DA4-A49C-29EF4A8369BA}{3BAE7E51-E3F4-41D0-853D9BB9FD47605F}]";
		string immediateScheduledTaskString = "";
		string immediateFileXMLString = "";
		string path = "";
		string author = "";
		string task_name = "";
		string command = "";
		string arguments = "";
		string version_string = "";
		string domain_string = "";
		string start = "";
		string end = "";
		string source_file_path = "";
		string dest_file = "";
		string dest_path = "";
		string dest_file_path = "";
		System.DirectoryServices.ActiveDirectory.Domain domain = null;
		int versionnumber = 0;
		int new_versionnumber = 0;
		int enable_flag = 0;
		int disable_flag = 3;
		
		
		try
		{
			if (argumentsize > 0)
			{
				if (args[0].Equals("help"))
				{
					Console.WriteLine(help);
				}
				else if (args[0].Equals("query"))
				{
					guid_filter = args[1];
					DirectoryEntry de = new DirectoryEntry();
					DirectorySearcher ds = new DirectorySearcher(de);
					ds.Filter = "(&(objectCategory=groupPolicyContainer)(name=" + guid_filter +"))";
					ds.PropertiesToLoad.Add("displayname");
					ds.PropertiesToLoad.Add("gpcMachineExtensionNames");
					ds.PageSize = 1000;
					ds.SizeLimit = 100;
					ds.PropertiesToLoad.Add("versionnumber");
					ds.PropertiesToLoad.Add("flags");
					ds.SearchScope = SearchScope.Subtree;
					SearchResultCollection src = ds.FindAll();
					string results = "";
					foreach (SearchResult sr in src)
					{
						ResultPropertyCollection myResultPropColl = sr.Properties;
						foreach (string myKey in myResultPropColl.PropertyNames)
						{
							foreach (Object myCollection in myResultPropColl[myKey])
							{
								results += myKey + ": " + myCollection + "\n\n";
							}
						}
					}
					Console.WriteLine(results);				
				}
				else if (args[0].Equals("itask"))
				{
					domain = System.DirectoryServices.ActiveDirectory.Domain.GetCurrentDomain();
					domain_string = domain.Name;
					guid_filter = args[1];
					guid_LDAP = args[2];
					gpcMacExtName = args[3];
					version_string = args[4];
					author = args[5];
					task_name = args[6];
					command = args[7];
					arguments = args[8];
					path = @"\\" + domain_string + "\\sysvol\\" + domain_string + "\\policies\\" + guid_filter;
					start =  @"<?xml version=""1.0"" encoding=""utf-8""?><ScheduledTasks
					clsid=""{CC63F200-7309-4ba0-B154-A71CD118DBCC}"">";
					end = @"</ScheduledTasks>";
					immediateScheduledTaskString = string.Format(@"<ImmediateTaskV2 clsid=""{{9756B581-76EC-4169-9AFC-0CA8D43ADB5F}}"" name=""{1}"" image=""0"" changed=""2019-03-30 23:04:20"" uid=""{4}""><Properties action=""C"" name=""{1}"" runAs=""NT AUTHORITY\System"" logonType=""S4U""><Task version=""1.3""><RegistrationInfo><Author>{0}</Author><Description></Description></RegistrationInfo><Principals><Principal id=""Author""><UserId>NT AUTHORITY\System</UserId><LogonType>S4U</LogonType><RunLevel>HighestAvailable</RunLevel></Principal></Principals><Settings><IdleSettings><Duration>PT10M</Duration><WaitTimeout>PT1H</WaitTimeout><StopOnIdleEnd>true</StopOnIdleEnd><RestartOnIdle>false</RestartOnIdle></IdleSettings><MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy><DisallowStartIfOnBatteries>true</DisallowStartIfOnBatteries><StopIfGoingOnBatteries>true</StopIfGoingOnBatteries><AllowHardTerminate>true</AllowHardTerminate><StartWhenAvailable>true</StartWhenAvailable><RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable><AllowStartOnDemand>true</AllowStartOnDemand><Enabled>true</Enabled><Hidden>false</Hidden><RunOnlyIfIdle>false</RunOnlyIfIdle><WakeToRun>false</WakeToRun><ExecutionTimeLimit>P3D</ExecutionTimeLimit><Priority>7</Priority><DeleteExpiredTaskAfter>PT0S</DeleteExpiredTaskAfter></Settings><Triggers><TimeTrigger><StartBoundary>%LocalTimeXmlEx%</StartBoundary><EndBoundary>%LocalTimeXmlEx%</EndBoundary><Enabled>true</Enabled></TimeTrigger></Triggers><Actions Context=""Author""><Exec><Command>{2}</Command><Arguments>{3}</Arguments></Exec></Actions></Task></Properties></ImmediateTaskV2>", author, task_name, command, arguments, Guid.NewGuid().ToString());
					versionnumber = Convert.ToInt32(version_string);
					new_gpcMacExtName = gpcMacExtName + schtask_gpcMacExtName;
					new_versionnumber = versionnumber + 5;
					if (Directory.Exists(path))
					{
						path += "\\Machine\\Preferences\\ScheduledTasks\\";
					}
					else
					{
						Console.Write("\n[");
						Console.BackgroundColor = ConsoleColor.Black;
						Console.ForegroundColor = ConsoleColor.Red;
						Console.Write("!");
						Console.ResetColor();
						Console.WriteLine("] Could not find the specified GPO path!");
					}
					if (!Directory.Exists(path))
					{
						System.IO.Directory.CreateDirectory(path);
					}
					path += "ScheduledTasks.xml";
					if (File.Exists(path))
					{
						Console.Write("\n[");
						Console.BackgroundColor = ConsoleColor.Black;
						Console.ForegroundColor = ConsoleColor.Red;
						Console.Write("!");
						Console.ResetColor();
						Console.WriteLine("] Warning, the GPO you are targetting already has a ScheduledTask.xml. At this time, download the XML at " + path + " and manually insert your task in for execution.");
					}
					else
					{
						Console.Write("\n[");
						Console.BackgroundColor = ConsoleColor.Black;
						Console.ForegroundColor = ConsoleColor.Green;
						Console.Write("+");
						Console.ResetColor();
						Console.WriteLine("] Creating file " + path);
						System.IO.File.WriteAllText(path, start + immediateScheduledTaskString + end);
						DirectoryEntry de = new DirectoryEntry(guid_LDAP);
						de.Properties["gpcMachineExtensionNames"].Clear();
						de.Properties["gpcMachineExtensionNames"].Add(new_gpcMacExtName);
						de.Properties["versionnumber"].Clear();
						de.Properties["versionnumber"].Add(new_versionnumber);
						de.CommitChanges();
						de.Close();
						DirectoryEntry de_check = new DirectoryEntry();
						DirectorySearcher ds = new DirectorySearcher(de_check);
						ds.Filter = "(&(objectCategory=groupPolicyContainer)(name=" + guid_filter +"))";
						ds.PageSize = 1000;
						ds.SizeLimit = 100;
						ds.PropertiesToLoad.Add("displayname");
						ds.PropertiesToLoad.Add("gpcMachineExtensionNames");
						ds.PropertiesToLoad.Add("versionnumber");
						ds.SearchScope = SearchScope.Subtree;
						SearchResultCollection src = ds.FindAll();
						string results = "";
						foreach (SearchResult sr in src)
						{
							ResultPropertyCollection myResultPropColl = sr.Properties;
							foreach (string myKey in myResultPropColl.PropertyNames)
							{
								foreach (Object myCollection in myResultPropColl[myKey])
								{
									results += myKey + ": " + myCollection + "\n";
								}
							}
						}
						Console.Write("[");
						Console.BackgroundColor = ConsoleColor.Black;
						Console.ForegroundColor = ConsoleColor.Green;
						Console.Write("+");
						Console.ResetColor();
						Console.WriteLine("] New Active Directory properties for GPO " + guid_filter + " are: \n\n" + results);
					}								
				}
				else if (args[0].Equals("file"))
				{
					domain = System.DirectoryServices.ActiveDirectory.Domain.GetCurrentDomain();
					domain_string = domain.Name;
					guid_filter = args[1];
					guid_LDAP = args[2];
					gpcMacExtName = args[3];
					version_string = args[4];
					source_file_path = args[5];
					dest_file = args[6];
					dest_path = args[7];
					dest_file_path = dest_path + "\\" + dest_file;
					path = @"\\" + domain_string + "\\sysvol\\" + domain_string + "\\policies\\" + guid_filter;
					start =  @"<?xml version=""1.0"" encoding=""utf-8""?><Files clsid=""{215B2E53-57CE-475c-80FE-9EEC14635851}"">";
					end = @"</Files>";
					immediateFileXMLString = string.Format(@"<File clsid=""{{50BE44C8-567A-4ED1-B1D0-9234FE1F38AF}}"" name=""{1}"" status=""{1}"" image=""0"" changed=""2019-03-30 23:04:20"" uid=""{4}""><Properties action=""C"" fromPath=""{3}"" targetPath=""{2}"" readOnly=""0"" archive=""0"" hidden=""1""/></File>", dest_file, dest_file, dest_file_path, source_file_path, Guid.NewGuid().ToString());
					versionnumber = Convert.ToInt32(version_string);
					new_gpcMacExtName = gpcMacExtName + file_gpcMacExtName;
					new_versionnumber = versionnumber + 5;
					if (Directory.Exists(path))
					{
						path += "\\Machine\\Preferences\\Files\\";
					}
					else
					{
						Console.Write("\n[");
						Console.BackgroundColor = ConsoleColor.Black;
						Console.ForegroundColor = ConsoleColor.Red;
						Console.Write("!");
						Console.ResetColor();
						Console.WriteLine("] Could not find the specified GPO path!");
					}
					if (!Directory.Exists(path))
					{
						System.IO.Directory.CreateDirectory(path);
					}
					path += "Files.xml";
					if (File.Exists(path))
					{
						Console.Write("\n[");
						Console.BackgroundColor = ConsoleColor.Black;
						Console.ForegroundColor = ConsoleColor.Red;
						Console.Write("!");
						Console.ResetColor();
						Console.WriteLine("] Warning, the GPO you are targetting already has a Files.xml. At this time, download the XML at " + path + " and manually insert your file in for creation.");
					}
					else
					{
						Console.Write("\n[");
						Console.BackgroundColor = ConsoleColor.Black;
						Console.ForegroundColor = ConsoleColor.Green;
						Console.Write("+");
						Console.ResetColor();
						Console.WriteLine("] Creating file " + path);
						System.IO.File.WriteAllText(path, start + immediateFileXMLString + end);
						DirectoryEntry de = new DirectoryEntry(guid_LDAP);
						de.Properties["gpcMachineExtensionNames"].Clear();
						de.Properties["gpcMachineExtensionNames"].Add(new_gpcMacExtName);
						de.Properties["versionnumber"].Clear();
						de.Properties["versionnumber"].Add(new_versionnumber);
						de.CommitChanges();
						de.Close();
						DirectoryEntry de_check = new DirectoryEntry();
						DirectorySearcher ds = new DirectorySearcher(de_check);
						ds.Filter = "(&(objectCategory=groupPolicyContainer)(name=" + guid_filter +"))";
						ds.PageSize = 1000;
						ds.SizeLimit = 100;
						ds.PropertiesToLoad.Add("displayname");
						ds.PropertiesToLoad.Add("gpcMachineExtensionNames");
						ds.PropertiesToLoad.Add("versionnumber");
						ds.SearchScope = SearchScope.Subtree;
						SearchResultCollection src = ds.FindAll();
						string results = "";
						foreach (SearchResult sr in src)
						{
							ResultPropertyCollection myResultPropColl = sr.Properties;
							foreach (string myKey in myResultPropColl.PropertyNames)
							{
								foreach (Object myCollection in myResultPropColl[myKey])
								{
									results += myKey + ": " + myCollection + "\n";
								}
							}
						}
						Console.Write("[");
						Console.BackgroundColor = ConsoleColor.Black;
						Console.ForegroundColor = ConsoleColor.Green;
						Console.Write("+");
						Console.ResetColor();
						Console.WriteLine("] New Active Directory properties for GPO " + guid_filter + " are: \n\n" + results);
					}								
				}
				else if (args[0].Equals("set"))
				{
					guid_filter = args[1];
					guid_LDAP = args[2];
					gpcMacExtName = args[3];
					version_string = args[4];
					versionnumber = Convert.ToInt32(version_string);
					DirectoryEntry de = new DirectoryEntry(guid_LDAP);
					de.Properties["gpcMachineExtensionNames"].Clear();
					de.Properties["gpcMachineExtensionNames"].Add(gpcMacExtName);
					de.Properties["versionnumber"].Clear();
					de.Properties["versionnumber"].Add(versionnumber);
					de.CommitChanges();
					de.Close();
					DirectoryEntry de_check = new DirectoryEntry();
					DirectorySearcher ds = new DirectorySearcher(de_check);
					ds.Filter = "(&(objectCategory=groupPolicyContainer)(name=" + guid_filter +"))";
					ds.PageSize = 1000;
					ds.SizeLimit = 100;
					ds.PropertiesToLoad.Add("displayname");
					ds.PropertiesToLoad.Add("gpcMachineExtensionNames");
					ds.PropertiesToLoad.Add("versionnumber");
					ds.SearchScope = SearchScope.Subtree;
					SearchResultCollection src = ds.FindAll();
					string results = "";
					foreach (SearchResult sr in src)
					{
						ResultPropertyCollection myResultPropColl = sr.Properties;
						foreach (string myKey in myResultPropColl.PropertyNames)
						{
							foreach (Object myCollection in myResultPropColl[myKey])
							{
								results += myKey + ": " + myCollection + "\n";
							}
						}
					}
					Console.Write("\n[");
					Console.BackgroundColor = ConsoleColor.Black;
					Console.ForegroundColor = ConsoleColor.Green;
					Console.Write("+");
					Console.ResetColor();
					Console.WriteLine("] New Active Directory properties for GPO " + guid_filter + " are: \n\n" + results);
				}
				else if (args[0].Equals("versionset"))
				{
					guid_filter = args[1];
					guid_LDAP = args[2];
					version_string = args[3];
					versionnumber = Convert.ToInt32(version_string);
					new_versionnumber = versionnumber + 5;
					DirectoryEntry de = new DirectoryEntry(guid_LDAP);
					de.Properties["versionnumber"].Clear();
					de.Properties["versionnumber"].Add(new_versionnumber);
					de.CommitChanges();
					de.Close();
					DirectoryEntry de_check = new DirectoryEntry();
					DirectorySearcher ds = new DirectorySearcher(de_check);
					ds.Filter = "(&(objectCategory=groupPolicyContainer)(name=" + guid_filter +"))";
					ds.PageSize = 1000;
					ds.SizeLimit = 100;
					ds.PropertiesToLoad.Add("displayname");
					ds.PropertiesToLoad.Add("versionnumber");
					ds.SearchScope = SearchScope.Subtree;
					SearchResultCollection src = ds.FindAll();
					string results = "";
					foreach (SearchResult sr in src)
					{
						ResultPropertyCollection myResultPropColl = sr.Properties;
						foreach (string myKey in myResultPropColl.PropertyNames)
						{
							foreach (Object myCollection in myResultPropColl[myKey])
							{
								results += myKey + ": " + myCollection + "\n";
							}
						}
					}
					Console.Write("\n[");
					Console.BackgroundColor = ConsoleColor.Black;
					Console.ForegroundColor = ConsoleColor.Green;
					Console.Write("+");
					Console.ResetColor();
					Console.WriteLine("] New Active Directory properties for GPO " + guid_filter + " are: \n\n" + results);
				}
				
				else if (args[0].Equals("enable"))
				{
					guid_filter = args[1];
					guid_LDAP = args[2];
					version_string = args[3];
					versionnumber = Convert.ToInt32(version_string);
					new_versionnumber = versionnumber + 5;
					DirectoryEntry de = new DirectoryEntry(guid_LDAP);
					de.Properties["versionnumber"].Clear();
					de.Properties["versionnumber"].Add(new_versionnumber);
					de.Properties["flags"].Clear();
					de.Properties["flags"].Add(enable_flag);
					de.CommitChanges();
					de.Close();
					DirectoryEntry de_check = new DirectoryEntry();
					DirectorySearcher ds = new DirectorySearcher(de_check);
					ds.Filter = "(&(objectCategory=groupPolicyContainer)(name=" + guid_filter +"))";
					ds.PageSize = 1000;
					ds.SizeLimit = 100;
					ds.PropertiesToLoad.Add("displayname");
					ds.PropertiesToLoad.Add("versionnumber");
					ds.PropertiesToLoad.Add("flags");
					ds.SearchScope = SearchScope.Subtree;
					SearchResultCollection src = ds.FindAll();
					string results = "";
					foreach (SearchResult sr in src)
					{
						ResultPropertyCollection myResultPropColl = sr.Properties;
						foreach (string myKey in myResultPropColl.PropertyNames)
						{
							foreach (Object myCollection in myResultPropColl[myKey])
							{
								results += myKey + ": " + myCollection + "\n";
							}
						}
					}
					Console.Write("\n[");
					Console.BackgroundColor = ConsoleColor.Black;
					Console.ForegroundColor = ConsoleColor.Green;
					Console.Write("+");
					Console.ResetColor();
					Console.WriteLine("] Success! New Active Directory properties for GPO " + guid_filter + " are: \n\n" + results);
				}
				
				else if (args[0].Equals("disable"))
				{
					guid_filter = args[1];
					guid_LDAP = args[2];
					version_string = args[3];
					versionnumber = Convert.ToInt32(version_string);
					new_versionnumber = versionnumber + 5;
					DirectoryEntry de = new DirectoryEntry(guid_LDAP);
					de.Properties["versionnumber"].Clear();
					de.Properties["versionnumber"].Add(new_versionnumber);
					de.Properties["flags"].Clear();
					de.Properties["flags"].Add(disable_flag);
					de.CommitChanges();
					de.Close();
					DirectoryEntry de_check = new DirectoryEntry();
					DirectorySearcher ds = new DirectorySearcher(de_check);
					ds.Filter = "(&(objectCategory=groupPolicyContainer)(name=" + guid_filter +"))";
					ds.PageSize = 1000;
					ds.SizeLimit = 100;
					ds.PropertiesToLoad.Add("displayname");
					ds.PropertiesToLoad.Add("versionnumber");
					ds.PropertiesToLoad.Add("flags");
					ds.SearchScope = SearchScope.Subtree;
					SearchResultCollection src = ds.FindAll();
					string results = "";
					foreach (SearchResult sr in src)
					{
						ResultPropertyCollection myResultPropColl = sr.Properties;
						foreach (string myKey in myResultPropColl.PropertyNames)
						{
							foreach (Object myCollection in myResultPropColl[myKey])
							{
								results += myKey + ": " + myCollection + "\n";
							}
						}
					}
					Console.Write("[");
					Console.BackgroundColor = ConsoleColor.Black;
					Console.ForegroundColor = ConsoleColor.Green;
					Console.Write("+");
					Console.ResetColor();
					Console.WriteLine("] Success! New Active Directory properties for GPO " + guid_filter + " are: \n\n" + results);
				}
					
				else
				{
					Console.WriteLine(help);
				}				
			}
			else
			{
				Console.Write("\n[");
				Console.BackgroundColor = ConsoleColor.Black;
				Console.ForegroundColor = ConsoleColor.Red;
				Console.Write("!");
				Console.ResetColor();
				Console.WriteLine("] Error: Not enough arguments. Type help to see usage.");
			}
		}
		catch (System.Runtime.InteropServices.COMException excep)
		{
			string error = excep.ToString();
			if (error.Contains("\nAccess is denied"))
			{
				Console.Write("\n[");
				Console.BackgroundColor = ConsoleColor.Black;
				Console.ForegroundColor = ConsoleColor.Red;
				Console.Write("!");
				Console.ResetColor();
				Console.WriteLine("] Access is denied");
			}
			else if (error.Contains("\nA referral was returned from the server"))
			{
				Console.Write("\n[");
				Console.BackgroundColor = ConsoleColor.Black;
				Console.ForegroundColor = ConsoleColor.Red;
				Console.Write("!");
				Console.ResetColor();
				Console.WriteLine("] Invalid LDAP path.");
			}
			else
			{
				Console.WriteLine(error);
			}
			System.Environment.Exit(-1);
		}
			
	}
}
