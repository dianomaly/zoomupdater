Description
=================

This script will query your zoom voice users for their phone numbers and will update their Actice Directory telephone number.

Usage
=================

1. Clone this repo unto a windows machine.
2. Add your Zoom Api token to the `$apitoken` param in the **zoomupdater.ps1** file.
3. Setup a scheduled task to run the **zoomupdater.ps1** script as a user that has access to update users in your Active Directory.

This script logs to the **`C:\logs\zoomupdater\zoomupdater.log`** location on the machine.

Requirements
=================

[New-Log4NetLogger] (https://github.com/swys/New-Log4NetLogger) 
[curl for windows] (https://curl.haxx.se/windows/)

A user with access to update users in Active Directory.

