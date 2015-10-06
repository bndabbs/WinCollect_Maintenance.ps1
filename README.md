WinCollect is a log forwarding agent used by IBM for their QRadar appliance. Unfortunately, this product is very unstable and has to be restarted every couple of days. 

This script will stop the service, gather logs that IBM wants for diagnostics, and start the service again.

I have written this specific to WinCollect, but it could be easily tailored to gather other log files and deliver them.

** Be sure to update lines 2 and 34-38 to match your environment.**
