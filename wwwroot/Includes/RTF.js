//Inda: Original RTF.js file was not Mozilla compatible
//Changed: Tags are now placed at the cursor position if the selection string is empty
//Added: Mozilla compatibility

function changeMozilla(sTag, bSingleTag, bSmiley)
{
 // Create attributes string
 var sAttributes = "";
	   
 for(var i=3; i<changeMozilla.arguments.length; i+=2) {
  sAttributes += " " + changeMozilla.arguments[i]
  sAttributes += "=\"" + changeMozilla.arguments[i+1] + "\""
 }
	
 // Set the textarea object
 var oTA = document.forms['AddEntry'].Content;
	
 // Get textarea value
 var s = document.forms['AddEntry'].Content.value;

 // Set new string to the text before the selection
 var sChanged = s.substring(0, oTA.selectionStart);
	
 // Add opening tag and attributes
 sChanged += (bSmiley) ? sTag : "<" + sTag + sAttributes + ">";
	
 // Add selected text
 sChanged += s.substring(oTA.selectionStart, oTA.selectionEnd);
	
 // Add closing tag
 if(!bSingleTag) sChanged += "</" + sTag + ">";
	
 // Add text after selection
 sChanged += s.substring(oTA.selectionEnd);
	
 // Change the textarea value
 oTA.value = sChanged;
}


function changeIE(sTag, bSingleTag, bSmiley)
{	
	//create attributes string
	var sAttributes = "";
	   
	for(var i=3; i<changeIE.arguments.length; i+=2)
	{
		sAttributes += " " + changeIE.arguments[i]
		sAttributes += "=\"" + changeIE.arguments[i+1] + "\""
	}
	
	
	//set the textarea object
	var oTA = document.forms['AddEntry'].Content;
	
	//to stop the button value changing if the textarea doesn't have the focus
	oTA.focus();
	
	//set selection text
	var s = document.selection.createRange().text;
	
	//create the changed text
	var sChanged = ""
	sChanged += (bSmiley) ? sTag + s : "<" + sTag + sAttributes + ">" + s;
	
	//add the closing tag
	if(!bSingleTag) sChanged += "</" + sTag + ">"
	
	//change the textarea text
	document.selection.createRange().text = sChanged;	
}


function boldThis()
{
	if(document.forms['AddEntry'].Content.selectionStart > -1) //Mozilla
	{
		//alert("Mozilla") //debug
		changeMozilla("b", false, false);
	}
	else if(document.selection && document.selection.createRange) //IE
	{
		//alert("IE") //debug
		changeIE("b", false, false);
	}
	else
	{
		//no support
		alert("Your browser is not supported");
	}
}


function italicsThis()
{
	if(document.forms['AddEntry'].Content.selectionStart > -1) //Mozilla
	{
		changeMozilla("i", false, false);
	}
	else if(document.selection && document.selection.createRange) //IE
	{
		changeIE("i", false, false);
	}
	else
	{
		alert("Your browser is not supported");
	}
}


function underlineThis()
{
	if(document.forms['AddEntry'].Content.selectionStart > -1) //Mozilla
	{
		changeMozilla("span", false, false, "style","text-decoration: underline");
	}
	else if(document.selection && document.selection.createRange) //IE
	{
		changeIE("span", false, false, "style","text-decoration: underline");
	}
	else
	{
		alert("Your browser is not supported");
	}
}


function crossThis()
{
	if(document.forms['AddEntry'].Content.selectionStart > -1) //Mozilla
	{
		changeMozilla("s", false, false);
	}
	else if(document.selection && document.selection.createRange) //IE
	{
		changeIE("s", false, false);
	}
	else
	{
		alert("Your browser is not supported");
	}
}


function leftThis()
{
	if(document.forms['AddEntry'].Content.selectionStart > -1) //Mozilla
	{
		changeMozilla("p", false, false, "align", "left");
	}
	else if(document.selection && document.selection.createRange) //IE
	{
		changeIE("p", false, false, "align", "left");
	}
	else
	{
		alert("Your browser is not supported");
	}
}


function centerThis()
{
	if(document.forms['AddEntry'].Content.selectionStart > -1) //Mozilla
	{
		changeMozilla("p", false, false, "align", "center");
	}
	else if(document.selection && document.selection.createRange) //IE
	{
		changeIE("p", false, false, "align", "center");
	}
	else
	{
		alert("Your browser is not supported");
	}
}


function rightThis()
{
	if(document.forms['AddEntry'].Content.selectionStart > -1) //Mozilla
	{
		changeMozilla("p", false, false, "align", "right");
	}
	else if(document.selection && document.selection.createRange) //IE
	{
		changeIE("p", false, false, "align", "right");
	}
	else
	{
		alert("Your browser is not supported");
	}
}


function lineThis()
{
	if(document.forms['AddEntry'].Content.selectionStart > -1) //Mozilla
	{
		changeMozilla("hr", true, false);
	}
	else if(document.selection && document.selection.createRange) //IE
	{
		changeIE("hr", true, false);
	}
	else
	{
		alert("Your browser is not supported");
	}
}

function linkThis()
{
	//test browser first so as not to annoy users
	
	if(document.forms['AddEntry'].Content.selectionStart > -1) //Mozilla
	{
		//prompt for URL
		var sURL = prompt("URL for the link.","http://")
		
		//test if cancel button was pressed
		if(sURL == null) return false;

		//after asking about a new window, call changeMozilla function to change the text
		if(confirm("Open link in a new window?"))
		{
			changeMozilla("a", false, false, "href", sURL, "target", "_New");
		}
		else
		{
			changeMozilla("a", false, false, "href", sURL);
		}
	}
	else if(document.selection && document.selection.createRange) //IE
	{
		var sURL = prompt("URL for the link.","http://")
	
		if(sURL == null) return false;
				
		if(confirm("Open link in a new window?"))
		{
			changeIE("a", false, false, "href", sURL, "target", "_New");
		}
		else
		{
			changeIE("a", false, false, "href", sURL);
		}
	}
	else
	{
		alert("Your browser is not supported");
	}
}

function AdvancedThis(sTag)
{

	//test if nothing was entered
	if(sTag == "") return false;

	if(document.forms['AddEntry'].Content.selectionStart > -1) //Mozilla
	{
		changeMozilla(sTag, true, false);
	}
	else if(document.selection && document.selection.createRange) //IE
	{
		changeIE(sTag, true, false);
	}
	else
	{
		alert("Your browser is not supported");
	}
}

function imageThis(from) {

 var left = (screen.width/2)-(400/2);
 var top = (screen.height/2)-(200/2);

 document.forms['AddEntry'].Content.focus();

 if (from == "") {
  popupWin = window.open('UploadPicture.asp','new_page','width=400,height=200,top='+top+',left='+left+',scrollbars=no')
 } else {
  popupWin = window.open('UploadPicture.asp?'+from,'new_page','width=400,height=200,top='+top+',left='+left+',scrollbars=no')
 }

}

function photoThis()
{
	if(document.forms['AddEntry'].Content.selectionStart > -1) //Mozilla
	{
		changeMozilla("p", false, false, "class", "dropshadow");
	}
	else if(document.selection && document.selection.createRange) //IE
	{
		changeIE("p", false, false, "class", "dropshadow");
	}
	else
	{
		alert("Your browser is not supported");
	}
}


//this is a beast of a function that will have to wait.
//test for mozilla and point the users towards http://spellbound.sourceforge.net/
function SpellThis()
{
	if(document.forms['AddEntry'].Content.selectionStart > -1) //Mozilla
	{
		alert("Your browser is not supported.\nIf you are using a Mozilla browser you could try using the SpellBound Extension.\nhttp://spellbound.sourceforge.net/")
	}
	else if(document.selection && document.selection.createRange) //IE
	{
		document.forms['AddEntry'].Content.select(); 
		document.forms['AddEntry'].Content.focus(); 
		Copied = document.forms['AddEntry'].Content.createTextRange();
		Copied.execCommand("RemoveFormat"); 
		Copied.execCommand("Copy");
		popupWin = window.open('Spell.asp','new_page','width=400,height=200,scrollbars=yes')
	}
	else
	{
		alert("Your browser is not supported");
	}
}

function show(id) {

    try {
	if(document.all.item) //IE
	 {
		var Advanceditem = document.all.item(id)

		if (Advanceditem != null)
		{
			if (Advanceditem.length != null)
			{
			    for (i=0; i<Advanceditem.length; i++) { Advanceditem(i).style.display = "inline"; }
			    document.getElementById('AdvancedButton').style.display = 'none';                                                                       
			}
		}
	 } else {
		//alert("Your browser is not supported");
	 }
	} catch (ex) {
	 // Annoyingly FireFox 2 requires this.
	}
}

function hide(id) {

    try {
	if(document.all) //IE
	 {
		var Advanceditem = document.all.item(id)
		
		if (Advanceditem != null)
		{
			if (Advanceditem.length != null)
			{
			   for (i=0; i<Advanceditem.length; i++) { Advanceditem(i).style.display = "none"; }
                        document.getElementById('AdvancedButton').style.display = 'inline';                                                                
			}
		}
	 } else {
		//alert("Your browser is not supported");
	 }
	} catch (ex) {
	 // Annoyingly FireFox 2 requires this.
	}
}

// Show/Hide functions for non-pointer layer/objects
window.onload=function(){
     hide('AdvancedTools');
}