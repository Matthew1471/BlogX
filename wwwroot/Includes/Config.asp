<!-- #INCLUDE FILE="Dim.asp" -->
<%
'***********************************************'
'Your site, database path & settings (CHANGE THESE!!!)
'***********************************************'

'-- Site settings --'
SiteURL = "http://blogx.co.uk/"
SharedFilesPath = "C:\Inetpub\wwwroot\Download"
Version = "2.2"

'-- Database settings --'
DataFile = "C:\Inetpub\Database\BlogX.mdb"
DataPassword = "DBPASS"

'-- Admin privileges / features (for hosted blogs) --'
AboutPage = True          'Use your About.asp file...or if false, link to the one on BlogX.co.uk instead.
AllowEditingLinks = 1	  'Setting to 1 allows admin to edit the links online.
ArgoSoftMailServer = 0    'If your server is running "ArgosoftMailServer" you can post from e-mail.
CalendarCheck = 1         'Checks whether all calendar dates have entries, Changing this to 0 increases BlogX's speed.

CommentNotify = 1                      'E-mails you when your blog is commented on.
NoEmailAddress = "noreply@youraddress.com" ' Use this address when you don't want to give away your real address.

LegacyMode = False 	  'Removes all functionality that was not in the ORIGINAL BlogX.
MailingList = 0           'Allows mailing list functionality.
'MultiLanguage = False     'Not Yet Implemented: Let your users change the engine to their language.
NotifyPingOMatic = 1      'Lists your Blog on PingOMatic.com whenever you make a post.
NoDate = 0		  'Stops separating entries by day.

'This displays an eye catching notice on ALL pages, empty by default.
NoticeText = ""


NoAdvertIP = "82.6.19.136" 'The address that is excluded from having adverts displayed.
OtherLinks = 1		   'Display the section "Other Links".
Register = True 	   'Tells BlogX.co.uk to add your site to the "WhoUses" page.
RSS = 1                    'Enables/Disables the RSS feed.
RSSImage = 1               'Displays the image (from the RSS folder) in your RSS feed.

SSLSupported = False      'Logins will be sent over your SSL connection.. You must have a valid certificate installed.

UseImagesInEditor = 1     'If 0, Buttons are used in the editor, otherwise if 1, Images are used in editor (e.g. Bold, Center, Underline)
UseExternalPlugin = 1     'If 1, Will dynamically include "Plugin.asp" in the navigation.

'-- Server time settings --'
ServerTimeOffset = 0 ' Corrects the server's time.
                     ' Enter how many minutes need to be added (add a negative number to subtract).
                     ' e.g. If this webserver reports the time as 1pm and it's actually 2pm in your time zone, set to 60 as 60 adds 60 minutes.
                     ' This only forces time correction, you may have to over-ride DST for the server's time zone in TimeZoneDetect.asp.

'***********************************************'
'END OF EDITING (DO NOT EDIT PAST THIS LINE)
'***********************************************'
%>
<!-- #INCLUDE FILE="Database.asp" -->