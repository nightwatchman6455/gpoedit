# gpoedit
A repository of files containing METRONOME as well as the template files needed for GPO abuse.
METRONOME
Written by Nightwatchman6455

Used to manipulate GPO Active Directory settings. See

USAGE:

To query for the existing properties of a GPO in Active Directory, use: metronome.exe query <Target GUID> 

To add the necessary pieces to add a scheduled task to a GPO, use: metronome.exe schtasks <Target GUID> <LDAP Path to GPO> <Original gpcMachineExtensionNames> <Original versionnumber>

To only increment the versionnumber field, use: metronome.exe versionset <Target GUID> <LDAP Path to GPO> <Original versionnumber>

To revert a GPO to its original properties or set the properties to different values, use: metronome.exe set <Target GUID> <LDAP Path to GPO> <Original gpcMachineExtensionNames> <Original versionnumber>

Files included in the repository:
- GptTmpl.inf: Use this template to add users to a local administrators group within an organizational unit. Use the versionset command above to increment the versionnumber attribute within Active Directory when doing it
- ScheduledTasks_Hostname.xml: Task template for GPO abuse that allows you to target based on NetBios hostname
- ScheduledTasks_IP.xml: Task template for GPO abuse that allows you to target based on IP address
  
