# GUI-Removed-User
This a Graphical Interface script that I wrote to expidate the removal of user account within an Exchange/Active Directory Enviorment. 

The idea and creation of the script was to have a single resource that would disable, set delegation if needed, remove Office 365 Licensing,
migrate the disabled user to a new OU Container and strip all memberships other than a pre-determined primary group. 

This script has been sanitized from my current enviroment.  You will need to fill in @yourdomain.com to match your current email syntax. You will also need to pull your Office 365 Licensing and verrify that you have correct login premissions to change the user license status. 

You can get the Office 365 License SKU with Get-MsolAccountSku once connected. This needs to be set on the $LicenseMatch on line 261. It is used to compare the user license to the one you are wanting to remove.  
