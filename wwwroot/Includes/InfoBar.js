//-- This is used to make an SNS call AND to display the bar on command --//
var userClicked = false;

// Simply sets the image to our SNS server //
function makeSNSRequest(webURL) {
 userClicked = true;
 document.getElementById('SNSFirer').onerror = noSNS;
 document.getElementById('SNSFirer').src = 'http://localhost:632/add.sns?' + webURL + 'SNS/';      
}

// Displays the info bar //
function noSNS() {
 if (userClicked) {
   if (document.layers) {
    document.layers['infobar'].display = 'block';
   } else if (document.all) {
    document.all['infobar'].style.display = 'block';
   } else if (document.getElementById) {
    document.getElementById('infobar').style.display = 'block';
   }
 }
}