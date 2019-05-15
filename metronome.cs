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

public class GPOedit
{
	public static void Main(string[] args)
	{
		int argumentsize = args.Length;
		string help = "To query for the existing properties of a GPO in Active Directory, use: \n - metronome.exe query <Target GUID> \n\nTo add the necessary pieces to add a scheduled task to a GPO, use:\n - metronome.exe schtasks <Target GUID> <LDAP Path to GPO> <Original gpcMachineExtensionNames> <Original versionnumber>\n\nTo revert a GPO to its original properties or set the properties to different values, use:\n - metronome.exe set <Target GUID> <LDAP Path to GPO> <Original gpcMachineExtensionNames> <Original versionnumber>\n\nTo only increment the versionnumber field, use:\n - metronome.exe versionset <Target GUID> <LDAP Path to GPO> <Original versionnumber>";
		string guid_filter = "";
		string guid_LDAP = "";
		string gpcMacExtName = "";
		string new_gpcMacExtName = "";
		string schtask_gpcMacExtName = "[{AADCED64-746C-4633-A97C-D61349046527}{CAB54552-DEEA-4691-817E-ED4A4D1AFC72}]";
		string version_string = "";
		int versionnumber = 0;
		int new_versionnumber = 0;
		
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
				else if (args[0].Equals("schtasks"))
				{
					guid_filter = args[1];
					guid_LDAP = args[2];
					gpcMacExtName = args[3];
					version_string = args[4];
					versionnumber = Convert.ToInt32(version_string);
					new_gpcMacExtName = gpcMacExtName + schtask_gpcMacExtName;
					new_versionnumber = versionnumber + 5;
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
					
					Console.WriteLine("New Active Directory properties for GPO " + guid_filter + " are: \n\n" + results);			
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
					
					Console.WriteLine("New Active Directory properties for GPO " + guid_filter + " are: \n\n" + results);
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
					
					Console.WriteLine("New Active Directory properties for GPO " + guid_filter + " are: \n\n" + results);
				}
					
				else
				{
					Console.WriteLine("help");
				}				
			}
			else
			{
				Console.WriteLine("Error: Not enough arguments. Type help to see usage.");
			}
		}
		catch (System.Runtime.InteropServices.COMException excep)
		{
			string error = excep.ToString();
			if (error.Contains("Access is denied"))
			{
				Console.WriteLine("Access is denied");
			}
			else if (error.Contains("A referral was returned from the server"))
			{
				Console.WriteLine("Invalid LDAP path.");
			}
			else
			{
				Console.WriteLine(error);
			}
			System.Environment.Exit(-1);
		}
			
	}
}