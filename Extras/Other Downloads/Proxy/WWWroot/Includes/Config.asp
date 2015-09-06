<!-- #INCLUDE FILE="Dim.asp" -->
<%
'***********************************************'
'Your DatabasePath & Settings (CHANGE THESE!!!)
'***********************************************'

'-- Note : To Ease confusion, set your AdminUsername and AdminPassword as the same as your Blogs' --'
AdminUsername   = "admin"
AdminPassword   = "letmein"
BlogURL = "http://theteenforum.co.uk/terst/"
SiteURL = "http://temp.matthew1471.co.uk/BlogXProxy/"

'-- Note : This is the credentials to the BlogXProxy DB NOT the BlogX DB --'
DataFile = "C:\Inetpub\Database\BlogXProxy.mdb"
DataPassword = "DBPASS"

LinksPath = "C:\Inetpub\Database\Links.txt"
OtherLinksPath = "C:\Inetpub\Database\OtherLinks.txt"

AboutPage         = True	       'If True, Will use the about page on your server.
BackgroundColor   = "#FFFFFF"	       'Should the stylesheet fail, this is the colour we fall back on.
Copyright         = "Matthew1471&copy;"   'The Text At The Bottom Of The Page
NotifyPingOMatic  = 1                  'Lists your Blog on the PingOMatic.com website whenever you make a post.
OtherLinks        = 1	               'Display the section "Other Links" from the OtherLinks.txt file.
SiteName          = "BlogX Proxy"      'Displayed at the top of the page
SiteDescription   = "Site Description" 'Again displayed at the top
SiteSubTitle      = "The BlogX Poster" 'Yet again
ShowCat           = True               'Relay a category too
Template          = "sandy"	       'The name of the folder (Inside "Templates") containing the CSS stylesheet.
UseImagesInEditor = 1                  'If 0, Buttons are used in the editor, otherwise if 1, Images are used in editor (e.g. Bold, Center, Underline)
UseExternalPlugin = 1                  'If 1, Will dynamically include "Plugin.asp" to the navigation
Version           = "1.0.00"	       'Used at the bottom of the page

CookieName        = "AdminSessionSecure2005" 'You should change this to prevent a clever hacker gaining access
UserCookieName    = "UserSessionSecure2005"  'Same, but for "User" access, THIS MUST NOT BE THE SAME VALUE AS ABOVE

'***********************************************'
'END OF EDITING (DON'T EDIT PAST THIS LINE)
'***********************************************'
%>
<!-- #INCLUDE FILE="Database.asp" -->