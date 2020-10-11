//for more explanation, visit https://www.squirrelwebdesign.com/2020/10/how-i-create-and-send-personalised-pdf-from-Google-Forms-for-free.html

function getResponse() {
  var sheet_app = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = sheet_app.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  
  var username = sheet.getRange(lastRow,lastColumn).getValue();
  var email = sheet.getRange(lastRow,2).getValue();
  var total = 0;
  for(var i=3;i<lastColumn;i++){
    var info = sheet.getRange(lastRow,i).getValue();
    if(info == 'Yes') {total++;}
  }
  
  var result_percent = total/(lastColumn-3)*100;
  
  sendHTML(username,result_percent,email);
}


function sendHTML(username, result_percent, email){
  
  var result_img = 'https://cdn.pixabay.com/photo/2015/11/09/11/53/morning-1035090_960_720.jpg';
  var result_desc = 'You are a true introvert! You are often perceived as quiet and reserved but that is just who you are! You tend to enjoy your own company and do not mind missing out a social gathering. But you do want to mix once in a while. Maybe just grab your nearest device and initiate a conversation with others.';
  
  //image and description if result is between 40 to 79
  if(result_percent > 39 && result_percent < 80){
    result_img = "https://cdn.pixabay.com/photo/2020/05/30/18/12/vintage-5239872_960_720.jpg";
    result_desc = "Wow, you do have some personality traits of an introvert. Although you are able to mix around in a social gathering, deep down, all you just want is to be alone with a mug of warm chocolate. All these gatherings are draining out your energy! Remember to do a social media detox once in a while to fully recharge your energy level!";
  //image and description if result is between 0 - 39
  }else if (result_percent < 40){
    result_img = "https://cdn.pixabay.com/photo/2018/10/10/21/39/mask-3738427_960_720.jpg";
    result_desc = "Are you sure you're an introvert? You sure sound like an extrovert to me!";
  }else{}
  
  //html file 
  var blob_img = UrlFetchApp.fetch(result_img).getBlob();
  
  //image needs to be under base64 encode, if not image will not be shown in the pdf file
  var b64 = blob_img.getContentType() + ';base64,'+ Utilities.base64Encode(blob_img.getBytes());
  
  var html = '<h3>Congratulations, ' + username + '!</h3>'+
    '<h1><strong>You are ' + result_percent + '% Introvert</strong></h1>'+
    '<p><strong><img src="data:'+ b64 + '"></strong></p>'+
    '<p>' + result_desc + '</p>'+
    '<p>I hope you enjoy knowing yourself better!</p>'+
    '<p>&nbsp;</p>'+
    '<hr />'+
    '<p>www.squirrelwebdesign.com | 2020</p>'+
    '<p>&nbsp;</p>';
  
  //convert html file to a pdf form
  var html_blob = Utilities.newBlob(html, "text/html", "result.html");
  var pdf = html_blob.getAs("application/pdf");
  DriveApp.createFile(pdf).setName("result.pdf");
 
  //send message with pdf attached to quiz-goers
  var message = 'Hi ' + username +', prepare to know your result? Check out the attached. Enjoy!';
  MailApp.sendEmail({
    to: email,
    subject: "The moment of truth",
    htmlBody: message, attachments: pdf
  });

}
