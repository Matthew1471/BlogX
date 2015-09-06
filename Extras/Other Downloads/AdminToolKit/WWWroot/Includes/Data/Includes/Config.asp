<!-- #INCLUDE FILE="Dim.asp" -->
<%
'***********************************************'
'Your DatabasePath & Settings (CHANGE THESE!!!)
'***********************************************'
%>
<!-- #INCLUDE FILE="DataFile.asp" -->
<%
DataPassword = "DBPASS"
LinksPath = "C:\Inetpub\Database\Links.txt"
OtherLinksPath = "C:\Inetpub\Database\OtherLinks.txt"
Version = "1.0.5.09"

AboutPage = True          'Use your About.asp file...or if false, link to the one on BlogX.co.uk instead
AllowEditingLinks = 0	  'Setting to 1 allows a logged in user to edit the links file online
ArgoSoftMailServer = 0    'If your server is running "ArgosoftMailServer" you can post from e-mail
CalendarCheck = 1         'Checks whether all calendar dates have entries, Changing this to 0 increases BlogX's speed
CommentNotify = 1         'E-mails you when your blog is commented on
MailingList = 0           'Allows you to run a mailinglist and displays the link at the bottom.
LegacyMode = False 	  'This will remove all functionality that wasn't in the ORIGINAL BlogX
NotifyPingOMatic = 1      'Lists your Blog on the PingOMatic.com website whenever you make a post.
NoDate = 0		  'Stops separating entries by day
OtherLinks = 1		  'Display the section "Other Links" from the OtherLinks.txt file
Register = True	  	  'Notifies BlogX.co.uk That You want Your Site Addded To OtherLinks.txt
RSS = 1                   'Lets people access your small RSS feed
RSSImage = 1              'Displays an image in your RSS feed (The image in the RSS folder).
TimeOffset = 0            'Change the time by this many hours e.g. "6" adds 6 hours to the server's time
UseImagesInEditor = 1     'If 0, Buttons are used in the editor, otherwise if 1, Images are used in editor (e.g. Bold, Center, Underline)
UseExternalPlugin = 0     'If 1, Will dynamically include "Plugin.asp" to the navigation
XMLTimeZone = "GMT"	  'The times specfied in the RSS are "GMT", "UT", "EST", "EDT", "CST", "CDT", "MST", "MDT", "PST", "PDT" see RFC822

'***********************************************'
'END OF EDITING (DON'T EDIT PAST THIS LINE)
'***********************************************'
%>
<!-- #INCLUDE FILE="Database.asp" -->