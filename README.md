# Office-2016-Packaging

This is the postinstall.sh file for my method of repackaging Office 2016. I've also included a sample uninstall.sh script for those who need that capability, and it should be compatible with both Casper and Munki.

More detailed instructions can be found here: http://www.richard-purves.com/?p=79

# RemoveOffice2011
This script follows the work by Richard Purves, but adds the removal of all Office 2011 components except Lync. All receipts are forgotten and applications removed - including Microsoft Auto Update and Microsoft Error Reporting.

You could you the script as a Munki-on-demand if you wish.
