# Office-2016-Packaging

This is the postinstall.sh file for my method of repackaging Office 2016. I've also included a sample uninstall.sh script for those who need that capability, and it should be compatible with both Casper and Munki.

More detailed instructions can be found here: http://www.richard-purves.com/?p=79

# RemoveOffice2011
This script follows the work by Richard Purves, but adds the removal of all Office 2011 components except Lync. All receipts are forgotten and applications removed - including Microsoft Auto Update and Microsoft Error Reporting.

You could you the script as a Munki-on-demand if you wish.

Here's an example to add to your pkginfo file:
```
<key>OnDemand</key>
<true/>
<key>minimum_munki_version</key>
<string>2.3.0</string>

<key>preinstall_alert</key>
<dict>
  <key>alert_detail</key>
  <string>This will uninstall Office 2011 from your computer, excluding MS Lync. Click Remove to proceed with uninstallation.</string>
  <key>alert_title</key>
  <string>Remove Office 2011?</string>
  <key>cancel_label</key>
  <string>Cancel</string>
  <key>ok_label</key>
  <string>Remove</string>
</dict>

```

# Makefile
Luggage Makefile is included. Edit to match your requirements. You (might) need to copy the RemoveOffice2011.sh to postinstall
