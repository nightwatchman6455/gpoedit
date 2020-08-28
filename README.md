# gpoedit
A repository of files containing METRONOME as well as the template files needed for GPO abuse.
METRONOME
Written by Nightwatchman6455

Used to manipulate GPO Active Directory settings. See http://nightwatchman.me/post/184884366363/gpo-abuse-and-you for more info.

USAGE:

To query for the existing properties of a GPO in Active Directory, use: metronome.exe query <Target GUID> 

To add the necessary pieces to add an immediate scheduled task to a GPO, use: metronome.exe itasks <Target GUID> <LDAP Path to GPO> <Original gpcMachineExtensionNames> <Original versionnumber> <task author> <task name> <command> <arguments>
  
To deploy a file using a GPO, use: metronome.exe file <Target GUID> <LDAP Path to GPO> <Original gpcMachineExtensionNames> <Original versionnumber> <source file and path> <destination file> <destination path (with no trailing \)

To only increment the versionnumber field, use: metronome.exe versionset <Target GUID> <LDAP Path to GPO> <Original versionnumber>

To revert a GPO to its original properties or set the properties to different values, use: metronome.exe set <Target GUID> <LDAP Path to GPO> <Original gpcMachineExtensionNames> <Original versionnumber>
  
To enable or disable a GPO, use: metronome.exe enable\disable <Target GUID> <LDAP Path to GPO> <Original versionnumber>

Files included in the repository:
- GptTmpl.inf: Use this template to add users to a local administrators group within an organizational unit. Use the versionset command above to increment the versionnumber attribute within Active Directory when doing it
- ScheduledTasks_Hostname.xml: Task template for GPO abuse that allows you to target based on NetBios hostname
- ScheduledTasks_IP.xml: Task template for GPO abuse that allows you to target based on IP address
  
