# GUI-Removed-User
This a Graphical Interface script that I wrote to expidate the removal of user account within an Exchange/Active Directory Enviorment. 

V7.2 adds some checks and balances to verify the account is correct, the delegation account exists.  There was also a change in how the script process to remove all the memberships of the said user.

The idea and creation of the script was to have a single resource that would disable, set delegation if needed, remove Office 365 Licensing,
migrate the disabled user to a new OU Container and strip all memberships other than a pre-determined primary group. 

This script has been sanitized from my current enviroment.  

You will need to update lines 33 thru 42 to match your current environment.

You can get the Office 365 License SKU with Get-MsolAccountSku once connected. This needs to be set on the $LicenseMatch on line 42. It is used to compare the user license to the one you are wanting to remove.  
