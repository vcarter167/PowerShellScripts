This is a collection of PowerShell Scripts to aid in the automation of some redundant user management tasks. 
The main purpose is to bulk add users and/or organization contacts to distribution lists or shared mailboxes. 
You will need to have Admin privelages in your tenant to execute these scripts. 
Also, you will need a CSV with at least two columns. 
One entitled "Name" and the other "Email"
These parameters are where the scripts will gather the data to locate the user/contact within your tenant's Active Directory. 
There may be error messages indicating that -Member parameter can't be found but the scripts are still able to manipulate the members in the distro lists/shared mailbox.
