<%
OPTION EXPLICIT
Response.ContentType = "text/vnd.wap.wml"%>
<?xml version='1.0'?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml"> 
<wml>
	<card id="Login" title="Login To Blog"><p>
		Username : <input type="text" name="Username"/><br/>
		Password : <input type="password" name="Password"/><br/>
                <br/>
		<anchor title="Login">Login
		<go href="AddNew.asp" method="post">
		<postfield name="Username" value="$(Username)"/>
		<postfield name="Password" value="$(Password)"/>
		</go>
		</anchor></p>
	</card>
</wml>