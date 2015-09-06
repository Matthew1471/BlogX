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
Due to limitations in the programming enviroment, BlogX does not currently support automatically running more than one plugin, to do this, it is suggested you edit Includes/Nav.asp to include multiple plugins by doing this :

*** Find (In Includes\NAV.asp) : ***

If (UseExternalPlugin = 1) AND (LegacyMode = False) AND (FSODisabled = False) Then
%>
<!-- #INCLUDE FILE="Plugin.asp" -->
<!-- <%=PluginTitle%> -->
<div class="section">
<h3 class="sectionTitle"><%=PluginTitle%></h3>
<%=PluginText%>
</div><br/>
<%
End If

*** Replace With : ***
 
If (UseExternalPlugin = 1) AND (LegacyMode = False) AND (FSODisabled = False) Then
%>
<!-- #INCLUDE FILE="Plugin.asp" -->
<!-- <%=PluginTitle%> -->
<div class="section">
<h3 class="sectionTitle"><%=PluginTitle%></h3>
<%=PluginText%>
</div><br/>

<!-- #INCLUDE FILE="Plugin2.asp" -->
<!-- <%=PluginTitle%> -->
<div class="section">
<h3 class="sectionTitle"><%=PluginTitle%></h3>
<%=PluginText%>
</div><br/>
<%
End If

Where Plugin2.asp is the name of your secondary plugin (when extracting from a Plugin ZIP, be careful, you probably do not want to overwrite your existing plugin).