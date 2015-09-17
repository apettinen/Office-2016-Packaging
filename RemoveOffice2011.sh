#!/bin/bash
# Uninstalls Office 2011, forgets the package and cleans the mess
# designed to be run as a Munki OnDemand (or standalone script)
# This script relies on the work of:
# Richard Purves:
# https://github.com/franton/Office-2016-Packaging/blob/master/postinstall.sh
# Rich Trouton:
# https://derflounder.wordpress.com/2015/08/05/creating-an-office-2016-15-12-3-installer/
# And on work by http://www.officeformachelp.com/ (the removal script for office 2011)
# Massive amounts of work also done by the MANY admins on Slack MacAdmin #microsoft-office too!

# Antti Pettinen
# TUTMac @ tut.fi
# 16.09.2015
# added setting telemetry reporting to false

# Set up log file, folder and function
LOGFOLDER="/usr/local/logs/"
LOG=$LOGFOLDER"/Office-2011-Uninstall.log"
error=0
O2016LICFILE="com.microsoft.office.licensingV2.plist"

if [ ! -d "$LOGFOLDER" ]; then
  /bin/mkdir -p $LOGFOLDER
fi

logme()
{
  # Check to see if function has been called correctly
  if [ -z "$1" ]; then
    echo $( date )" - logme function call error: no text passed to function! Please recheck code!"
    exit 1
  fi

  # Log the passed details
  echo $( date )" - "$1 >> $LOG
  echo "" >> $LOG
}

# Office 2011 Install location

office2011="/Applications/Microsoft Office 2011"
logme "Checking if Office 2011 is present"

# If installed, then clean up files

if [ -d "$office2011" ]; then
  logme "Office 2011 installation detected. Removing."

  # Stop Office 2011 background processes
  logme "Stopping Office 2011 background processes"
  /usr/bin/osascript -e 'tell application "Microsoft Database Daemon" to quit' | tee -a ${LOG}
  /usr/bin/osascript -e 'tell application "Microsoft AU Daemon" to quit' | tee -a ${LOG}
  /usr/bin/osascript -e 'tell application "Office365Service" to quit' | tee -a ${LOG}

  # Delete external applications apart from Lync
  logme "Deleting Office 2011 applications"
  /bin/rm -R '/Applications/Microsoft Communicator.app/' | tee -a ${LOG}
  /bin/rm -R '/Applications/Microsoft Messenger.app/' | tee -a ${LOG}
  /bin/rm -R '/Applications/Microsoft Office 2011/' | tee -a ${LOG}
  /bin/rm -R '/Applications/Remote Desktop Connection.app/' | tee -a ${LOG}

  # Delete Microsoft Auto Update
  if [ -d '/Library/Application Support/Microsoft/MAU2.0' ]; then
    logme "Deleting /Library/Application Support/Microsoft/MAU2.0"
    /bin/rm -R '/Library/Application Support/Microsoft/MAU2.0' | tee -a ${LOG}
  fi
  # Delete Microsoft Error Reporting
  if [ -d '/Library/Application Support/Microsoft/MERP2.0' ]; then
    logme "Deleting /Library/Application Support/Microsoft/MERP2.0"
    /bin/rm -R '/Library/Application Support/Microsoft/MERP2.0' | tee -a ${LOG}
  fi

  # Remove all Automator actions
  logme "Deleting Automator actions"
  /bin/rm -R /Library/Automator/*Excel* | tee -a ${LOG}
  /bin/rm -R /Library/Automator/*Office* | tee -a ${LOG}
  /bin/rm -R /Library/Automator/*Outlook* | tee -a ${LOG}
  /bin/rm -R /Library/Automator/*PowerPoint* | tee -a ${LOG}
  /bin/rm -R /Library/Automator/*Word* | tee -a ${LOG}
  /bin/rm -R /Library/Automator/*Workbook* | tee -a ${LOG}
  /bin/rm -R '/Library/Automator/Get Parent Presentations of Slides.action' | tee -a ${LOG}
  /bin/rm -R '/Library/Automator/Set Document Settings.action' | tee -a ${LOG}

  # Remove Office Fonts and copy disabled ones back into place
  logme "Deleting Microsoft Fonts folder"
  /bin/rm -R /Library/Fonts/Microsoft/ | tee -a ${LOG}

  logme "Moving previously disabled fonts back to main fonts folder"
  /bin/mv '/Library/Fonts Disabled/Arial Bold Italic.ttf' /Library/Fonts | tee -a ${LOG}
  /bin/mv '/Library/Fonts Disabled/Arial Bold.ttf' /Library/Fonts | tee -a ${LOG}
  /bin/mv '/Library/Fonts Disabled/Arial Italic.ttf' /Library/Fonts | tee -a ${LOG}
  /bin/mv '/Library/Fonts Disabled/Arial.ttf' /Library/Fonts | tee -a ${LOG}
  /bin/mv '/Library/Fonts Disabled/Brush Script.ttf' /Library/Fonts | tee -a ${LOG}
  /bin/mv '/Library/Fonts Disabled/Times New Roman Bold Italic.ttf' /Library/Fonts | tee -a ${LOG}
  /bin/mv '/Library/Fonts Disabled/Times New Roman Bold.ttf' /Library/Fonts | tee -a ${LOG}
  /bin/mv '/Library/Fonts Disabled/Times New Roman Italic.ttf' /Library/Fonts | tee -a ${LOG}
  /bin/mv '/Library/Fonts Disabled/Times New Roman.ttf' /Library/Fonts | tee -a ${LOG}
  /bin/mv '/Library/Fonts Disabled/Verdana Bold Italic.ttf' /Library/Fonts | tee -a ${LOG}
  /bin/mv '/Library/Fonts Disabled/Verdana Bold.ttf' /Library/Fonts | tee -a ${LOG}
  /bin/mv '/Library/Fonts Disabled/Verdana Italic.ttf' /Library/Fonts | tee -a ${LOG}
  /bin/mv '/Library/Fonts Disabled/Verdana.ttf' /Library/Fonts | tee -a ${LOG}
  /bin/mv '/Library/Fonts Disabled/Wingdings 2.ttf' /Library/Fonts | tee -a ${LOG}
  /bin/mv '/Library/Fonts Disabled/Wingdings 3.ttf' /Library/Fonts | tee -a ${LOG}
  /bin/mv '/Library/Fonts Disabled/Wingdings.ttf' /Library/Fonts | tee -a ${LOG}

  # Remove Sharepoint plugin
  logme "Deleting Sharepoint folder"
  /bin/rm -R /Library/Internet\ Plug-Ins/SharePoint* | tee -a ${LOG}

  # Remove LaunchDaemons, preference files and any helper tools
  logme "Deleting LaunchDaemons, Prefs and helper tools"
  /bin/rm -R /Library/LaunchDaemons/com.microsoft.* | tee -a ${LOG}
  #preserve office 2016 license file, all other preferences are removed
  # if you want to preserve other files, add the notation -not -name "preserve.this.file" before the -delete flag
  /usr/bin/find /Library/Preferences/ -type f -name "com.microsoft.*" -not -name "$O2016LICFILE" -delete  | tee -a ${LOG}
  /bin/rm -R /Library/PrivilegedHelperTools/com.microsoft.* | tee -a ${LOG}

  # Clean the receipt database i.e. forget packages:
  # Office 2016 applications are of form com.microsoft.package.ApplicationName
  # IF they are installed as standalone, MAS-style packages

  # Forget all Office receipts:
  logme "Forgeting package recipts for Office 2011"
  OFFICERECEIPTS=$(/usr/sbin/pkgutil --pkgs=com.microsoft.office.*)
	for ORCPT in $OFFICERECEIPTS; do
		/usr/sbin/pkgutil --forget $ORCPT
	done

  # Forget the Remote Desktop Connection receipts:
  # Microsoft Remote Desktop is com.microsoft.rdc.mac
  logme "Forgeting package receipts for Office 2011 Remote Desktop Connection"
  RDCRECEIPTS=$(/usr/sbin/pkgutil --pkgs=com.microsoft.rdc.all.*)
  for RRCPT in $RDCRECEIPTS; do
    /usr/sbin/pkgutil --forget $RRCPT
  done

  # Forget Messenger receipts
  logme "Forgeting package receipts for Office 2011 Messenger"
  MSGRECEIPTS=$(/usr/sbin/pkgutil --pkgs=com.microsoft.msgr.*)
  for MSGRCTP in $MSGRECEIPTS; do
    /usr/sbin/pkgutil --forget $MSGRCTP
  done

  # Forget Microsoft Auto Update receipts
  logme "Forgeting package receipts for Microsoft Auto Update"
  MAURECEIPTS=$(/usr/sbin/pkgutil --pkgs=com.microsoft.mau.*)
  for MAURCTP in $MAURECEIPTS; do
    /usr/sbin/pkgutil --forget $MAURCTP
  done

  # Forget Microsoft Error Reporting receipts
  logme "Forgeting package receipts for Microsoft Error Reporting"
  MERPRECEIPTS=$(/usr/sbin/pkgutil --pkgs=com.microsoft.merp.*)
  for MERPRCTP in $MERPRECEIPTS; do
    /usr/sbin/pkgutil --forget $MERPRCTP
  done

  # detect Office 2016 application installations and disable first-run dialogs
  # as they were removed by Office2011 uninstallation process:
  # NOTE! this will assume you have been using the "official" Microsoft receipts in Office 2016 installation
  # change them according to your receipts if necessary!

  # Excel 2016
  if /usr/sbin/pkgutil --pkgs=com.microsoft.package.Microsoft_Excel.app; then
    logme "Disabling Excel 2016 first-run dialog and setting telemetry reporting to MS to false"
    /usr/bin/defaults write $3/Library/Preferences/com.microsoft.Excel kSubUIAppCompletedFirstRunSetup1507 -bool true
    /usr/bin/defaults write $3/Library/Preferences/com.microsoft.Excel SendAllTelemetryEnabled -bool false
  fi

  # OneNote 2016
  if /usr/sbin/pkgutil --pkgs=com.microsoft.package.Microsoft_OneNote.app; then
    logme "Disabling OneNote 2016 first-run dialog and setting telemetry reporting to MS to false"
    /usr/bin/defaults write $3/Library/Preferences/com.microsoft.onenote.mac kSubUIAppCompletedFirstRunSetup1507 -bool true
    /usr/bin/defaults write $3/Library/Preferences/com.microsoft.onenote.mac SendAllTelemetryEnabled -bool false
  fi

  # Outlook 2016
  if /usr/sbin/pkgutil --pkgs=com.microsoft.package.Microsoft_Outlook.app; then
    logme "Disabling Outlook 2016 first-run dialog and setting telemetry reporting to MS to false"
    /usr/bin/defaults write $3/Library/Preferences/com.microsoft.Outlook kSubUIAppCompletedFirstRunSetup1507 -bool true
    /usr/bin/defaults write $3/Library/Preferences/com.microsoft.Outlook FirstRunExperienceCompletedO15 -bool true
    /usr/bin/defaults write $3/Library/Preferences/com.microsoft.Outlook SendAllTelemetryEnabled -bool false

  fi

  # PowerPoint 2016
  if /usr/sbin/pkgutil --pkgs=com.microsoft.package.Microsoft_PowerPoint.app; then
    logme "Disabling PowerPoint 2016 first-run dialog and setting telemetry reporting to MS to false"
    /usr/bin/defaults write $3/Library/Preferences/com.microsoft.PowerPoint kSubUIAppCompletedFirstRunSetup1507 -bool true
    /usr/bin/defaults write $3/Library/Preferences/com.microsoft.PowerPoint SendAllTelemetryEnabled -bool false
  fi

  # Word 2016
  if /usr/sbin/pkgutil --pkgs=com.microsoft.package.Microsoft_Word.app; then
    logme "Disabling Word 2016 first-run dialog and setting telemetry reporting to MS to false"
    /usr/bin/defaults write $3/Library/Preferences/com.microsoft.Word kSubUIAppCompletedFirstRunSetup1507 -bool true
    /usr/bin/defaults write $3/Library/Preferences/com.microsoft.Word SendAllTelemetryEnabled -bool false
  fi

else
  logme "Office 2011 not installed. Skipping uninstallation."
fi

logme "Office 2011 removal completed"
exit 0
