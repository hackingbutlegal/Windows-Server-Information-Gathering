These are a couple PowerShell scripts that use WMI to output Windows network hardware and software characteristics and output neatly to Excel for analysis.

If you are in charge of a network comprised of Windows servers, this can come in handy. I wrote these one day at work to help conduct information assurance audits.

Populate c:\Servers.txt with one (resolvable, or will timeout) hostname or ip per line.

You must open PowerShell with a network account that will have Administrator rights on the machines you intend to query. I prefer to run cmd.exe and then open PowerShell from that command line (and keep it open).

You may also need to login to the individual machines and install WMI (type "wmic" at the command line).


Happy recon!
