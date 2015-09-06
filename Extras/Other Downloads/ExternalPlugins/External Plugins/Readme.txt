-------------------------------
BlogX Intelligent Plugin System
-------------------------------
This folder is a folder for you to choose some plugins

BlogX now supports a thirdparty file "Plugin.asp" simply set "UseExternalPlugin" to equal "1" in the "Includes/Config.asp" and copy across a file to "Includes/Plugin.asp"..

----------------
Aquiring Plugins
----------------
You can aquire plugins from "BlogX.co.uk" or from any other site.

However, For your convienience, sample plugins are in this folder

--------------
Making Plugins
--------------
To make a plugin, simply program an ASP page, that stores the title and text in the variables

"PluginTitle" and "PluginText"

---------------------
Running more than one
---------------------
BlogX does not currently automatically support running more than one plugin (Due to limitations in the programming enviroment), if this is a MAJOR limitation to you, it is suggested you hack Nav.asp to include the plugin by doing this :

*** Find (In Includes\NAV.asp) : ***

If (UseExternalPlugin = 1) AND (LegacyMode = False) Then
%>
<!-- #INCLUDE FILE="Plugin.asp" -->
<!--- <%=PluginTitle%> --->
<DIV class=section>
<H3><%=PluginTitle%></H3>
<UL><%=PluginText%></UL>
</DIV><BR>
<%
End If

*** Replace With : ***
 
If (UseExternalPlugin = 1) AND (LegacyMode = False) Then
%>
<!-- #INCLUDE FILE="Plugin.asp" -->
<!--- <%=PluginTitle%> --->
<DIV class=section>
<H3><%=PluginTitle%></H3>
<UL><%=PluginText%></UL>
</DIV><BR>

<!-- #INCLUDE FILE="Plugin2.asp" -->
<!--- <%=PluginTitle%> --->
<DIV class=section>
<H3><%=PluginTitle%></H3>
<UL><%=PluginText%></UL>
</DIV><BR>

<%
End If

Where Plugin2.asp is the name of your secondary plugin (When extracting from a Plugin ZIP, Be careful, you probably don't want to overwrite your existing Plugin)