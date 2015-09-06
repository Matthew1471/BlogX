function boldThis(from) {strSelection = document.selection.createRange().text
document.AddEntry.Content.focus();
if (strSelection == "") {document.AddEntry.Content.value += "<b></b>"} 
else document.selection.createRange().text = "<b>" + strSelection + "</b>"}

function italicsThis(from) {strSelection = document.selection.createRange().text
document.AddEntry.Content.focus();
if (strSelection == "") {document.AddEntry.Content.value += "<i></i>"} 
else document.selection.createRange().text = "<i>" + strSelection + "</i>"}

function underlineThis(from) {strSelection = document.selection.createRange().text
document.AddEntry.Content.focus();
if (strSelection == "") {document.AddEntry.Content.value += "<u></u>"} 
else document.selection.createRange().text = "<u>" + strSelection + "</u>"}

function crossThis(from) {strSelection = document.selection.createRange().text
document.AddEntry.Content.focus();
if (strSelection == "") {document.AddEntry.Content.value += "<s></s>"} 
else document.selection.createRange().text = "<s>" + strSelection + "</s>"}

function leftThis(from) {strSelection = document.selection.createRange().text
document.AddEntry.Content.focus();
if (strSelection == "") {document.AddEntry.Content.value += "<p align=\"Left\"></p>"} 
else document.selection.createRange().text = "<p align=\"Left\">" + strSelection + "</p>"}

function centerThis(from) {strSelection = document.selection.createRange().text
document.AddEntry.Content.focus();
if (strSelection == "") {document.AddEntry.Content.value += "<p align=\"Center\"></p>"} 
else document.selection.createRange().text = "<p align=\"Center\">" + strSelection + "</p>"}

function rightThis(from) {strSelection = document.selection.createRange().text
document.AddEntry.Content.focus();
if (strSelection == "") {document.AddEntry.Content.value += "<p align=\"Right\"></p>"} 
else document.selection.createRange().text = "<p align=\"Right\">" + strSelection + "</p>"}

function lineThis(from) {document.AddEntry.Content.focus(); document.AddEntry.Content.value += " <hr>"}

function linkThis(from) {
			 document.AddEntry.Content.focus();
			 strSelection = document.selection.createRange().text
			 txt=prompt("URL for the link.","http://");
			 var newWind=confirm("Open link in a new window?");

		if (txt!=null) {


			if (strSelection == "") {
						 document.AddEntry.Content.value += " <a href=\""+txt+"\"";
						 if (newWind == true) {document.AddEntry.Content.value += " target=\"_New\"";}
						 document.AddEntry.Content.value += ">"+txt+"</a>";
				         } else {
						 document.selection.createRange().text = " <a href=\""+txt+"\"";
						 if (newWind == true) {document.selection.createRange().text += " target=\"_New\"";}
						 document.selection.createRange().text += ">" + strSelection + "</a>";}

					        }


				}

function imageThis(from) {document.AddEntry.Content.focus();
if (from == "") {popupWin = window.open('UploadPicture.asp','new_page','width=400,height=200,scrollbars=no')} 
else popupWin = window.open('UploadPicture.asp?'+from,'new_page','width=400,height=200,scrollbars=no')}

function photoThis(from) {document.AddEntry.Content.focus();
strSelection = document.selection.createRange().text
if (strSelection == "") {document.AddEntry.Content.value += "<p class=\"dropshadow\"></p>"} 
else document.selection.createRange().text = "<p class=\"dropshadow\">" + strSelection + "</p>"}

function SpellThis() {
document.AddEntry.Content.select(); 
document.AddEntry.Content.focus(); 
Copied = document.AddEntry.Content.createTextRange();
Copied.execCommand("RemoveFormat"); 
Copied.execCommand("Copy");
popupWin = window.open('Spell.asp','new_page','width=400,height=200,scrollbars=yes')
}



