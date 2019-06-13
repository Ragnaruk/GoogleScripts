function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  
  menu.addItem('Подготовить таблицу к работе', 'initializeActiveSheet');
  menu.addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function getVkToken() {
  var authorizationUrl = 'https://oauth.vk.com/authorize?client_id=6947304&display=page&redirect_uri=https://oauth.vk.com/blank.html&scope=video,offline&response_type=token&v=5.95&state=123456';
  
  var template = HtmlService.createTemplate('<style type="text/css">' +
                                            '	button#oauth {' +
                                            '		padding: 7px 16px 8px;' +
                                            '		margin: 0;' +
                                            '		font-size: 12.5px;' +
                                            '		display: inline-block;' +
                                            '		zoom: 1;' +
                                            '		cursor: pointer;' +
                                            '		white-space: nowrap;' +
                                            '		outline: none;' +
                                            '		font-family: -apple-system, BlinkMacSystemFont, Roboto, Helvetica Neue, sans-serif;' +
                                            '		vertical-align: top;' +
                                            '		line-height: 15px;' +
                                            '		text-align: center;' +
                                            '		text-decoration: none;' +
                                            '		background: none;' +
                                            '		background-color: #5181b8;' +
                                            '		color: #fff;' +
                                            '		border: 0;' +
                                            '		border-radius: 4px;' +
                                            '		box-sizing: border-box;' +
                                            '		width: 100%;' +
                                            '	}' +
                                            '</style>' +
                                            '<script>' +
                                            '	function redirect() {' +
                                            '		window.open("<?= authorizationUrl ?>", "_blank");' +
                                            '	}' +
                                            '</script>' +
                                            '<p>' +
                                            'После нажатия на кнопку ниже браузер перенаправит вас на сайт ВК, где нужно будет авторизоваться и разрешить приложению доступ к вашим данным.' +
                                            '</p>' +
                                            '<p>' +
                                            'После окончания авторизации вас перенаправит на страницу с надписью: "Пожалуйста, не копируйте данные из адресной строки для сторонних сайтов. Таким образом Вы можете потерять доступ к Вашему аккаунту."' +
                                            '</p>' +
                                            '<p>' +
                                            'Вам нужно будет скопировать адрес этой страницы из адресной строки в ячейку B1.' +
                                            '</p>' +
                                            '<button id="oauth" onclick="redirect()">Авторизоваться в ВК</button>');
  template.authorizationUrl = authorizationUrl;
  var page = template.evaluate();
  
  SpreadsheetApp.getUi().showSidebar(page);
}

var sheetActive = SpreadsheetApp.getActiveSpreadsheet();;
var sheetOptions = sheetActive.getSheets()[0];
var sheetComments = sheetActive.getSheetByName('Комментарии');;

function initializeActiveSheet() {
  // Delete all triggers
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  sheetOptions = sheetActive.getSheets()[0];
  sheetOptions.setName('Настройки')
  
  sheetActive.setActiveSheet(sheetOptions);
  
  sheetOptions.getRange('A1').setValue('Ссылка авторизации');
  sheetOptions.getRange('A2').setValue('Прямая ссылка на трансляцию');
  sheetOptions.getRange('A3').setValue('Дата и время окончания');
  
  sheetOptions.getRange('B3').setValue(new Date());
  sheetOptions.getRange('B3').setNumberFormat("dd.MM.yyyy hh:mm:ss");
  
  sheetOptions.getRange('B1').setNote('После запуска скрипта в правой части экрана появится ссылка на авторизацию и предоставление прав в ВК.' +
                                      ' При успешной авторизации произойдет перенаправление на страницу с адресом формата:' +
                                      ' https://oauth.vk.com/blank.html#access_token=XXXXX&expires_in=0&user_id=39199554&state=123456.' +
                                      ' Ссылку на эту страницу нужно вставить в это поле.');
  sheetOptions.getRange('B2').setNote('URL трансляции в формате: https://vk.com/videoXXXXX_XXXXX. Его можно получить под видео: Поделиться -> Экспортировать -> Прямая ссылка.');
  sheetOptions.getRange('B3').setNote('Скрипт будет выполняться до тех пор, пока не наступит время, указанное в ячейке.');
  
  sheetOptions.getRange('A1:A3').setFontWeight('bold');
  sheetOptions.getRange('A1:B3').setBorder(true, true, true, true, true, true);
  
  sheetOptions.autoResizeColumns(1, 2);
  
  // Show a VK auth sidebar
  getVkToken();
  
  // Create an onEdit trigger that will launch the main logic
  ScriptApp.newTrigger('onSheetEdit')
  .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
  .onEdit()
  .create();
}

// Function that returns true if script should continue running
function stopAllRunningScripts() {
  PropertiesService.getScriptProperties().deleteAllProperties();
}

// Launches every time sheet is edited
function onSheetEdit() {
  var cells = sheetOptions.getRange('B1:B3').getValues();
  
  var isReady = true;
  
  // Check whether all fields are filled
  for (var i = 0; i < cells.length && isReady; i++) {
    if (cells[i].toString().length === 0) {
      isReady = false;
    }
  }
  
  // If all fields are filled and time in B3 is later than now
  if (isReady && sheetOptions.getRange('B3').getValue().valueOf() > (new Date()).valueOf()) {
    // Disable all running scripts
    stopAllRunningScripts();
    
    // Parse value fields
    var userToken = sheetOptions.getRange('B1').getValue().toString();
    userToken = userToken.substring(userToken.search('#access_token=') + '#access_token='.length, userToken.search('&expires_in'));
    
    var videoUrl = sheetOptions.getRange('B2').getValue().toString();
    videoUrl = videoUrl.substring(videoUrl.search('https://vk.com/video') + 'https://vk.com/video'.length).split('_');
    var ownerId = videoUrl[0];
    var videoId = videoUrl[1];
    
    var offset = 0;
    var lineNumberOnSheet = 2;
  
    // Create an empty comments sheet and prepare it for writing
    if (sheetComments != null) {
      sheetActive.deleteSheet(sheetComments);
    }
    
    sheetComments = sheetActive.insertSheet();
    sheetComments.setName('Комментарии');
    
    sheetComments.getRange('A1').setValue('Имя пользователя');
    sheetComments.getRange('B1').setValue('Время комментария');
    sheetComments.getRange('C1').setValue('Текст комментария');
    sheetComments.getRange('D1').setValue('Ссылка на аватар');
    
    sheetComments.getRange('A1:D1').setFontWeight('bold');
    sheetComments.autoResizeColumns(1, 4);
    
    sheetComments.getRange('B2:B').setNumberFormat("dd.MM.yyyy hh:mm:ss");
    
    // Set global run property for a kill-switch
    var scriptId = (new Date).valueOf();
    PropertiesService.getScriptProperties().setProperty(scriptId, "running");
    
    // While current date is less than date in the cell and while global run property is true
    while (sheetOptions.getRange('B3').getValue().valueOf() > (new Date()).valueOf() && PropertiesService.getScriptProperties().getProperty(scriptId)) {
      var resp = receiveVKComments(userToken, ownerId, videoId, offset, lineNumberOnSheet);
      offset = resp[0];
      lineNumberOnSheet = resp[1];
      
      Utilities.sleep(1000);
    }
  }
}

function receiveVKComments(userToken, ownerId, videoId, offset, lineNumberOnSheet) {
  var url = 'https://api.vk.com/method/video.getComments?count=100&sort=asc&owner_id=' + ownerId + '&video_id=' + videoId + '&offset=' + 0 + '&access_token=' + userToken + '&v=5.95';
  
  var response = JSON.parse(UrlFetchApp.fetch(url).getContentText()).response;
  
  var numberOfComments = response.count;
  
  if (numberOfComments - offset > 30) {
    offset = numberOfComments - 30;
  }
  
  for (var i = offset; i < numberOfComments; i += 100) {
    // Time the speed of completion to avoid the limit of 3 requests / second
    var timeBegin = Date.now()
    
    if (numberOfComments - i < 100) {
      var count = numberOfComments - i;
    } else {
      var count = 100;
    }
    
    var url = 'https://api.vk.com/method/video.getComments?count=100&sort=asc&owner_id=' + ownerId + '&video_id=' + videoId + '&offset=' + i + '&access_token=' + userToken + '&v=5.95';
    var response = JSON.parse(UrlFetchApp.fetch(url).getContentText()).response;

    // Add all user ids to a string and separate them by commas
    var user_ids = "";
    for (var j = 0; j < response.count; j++) {
      if (response.items[j]) {
        user_ids += response.items[j].from_id + ",";
      }
    }
    user_ids = user_ids.slice(0, -1);
    
    var urlUser = 'https://api.vk.com/method/users.get?user_ids=' + user_ids + '&fields=photo_50&access_token=' + userToken + '&v=5.95';
    var responseUser = JSON.parse(UrlFetchApp.fetch(urlUser).getContentText()).response;
    
    var userNames = {};    
    var userAvatars = {};
    
    for (var j = 0; j < responseUser.length; j++) {
      var id = responseUser[j].id;
      
      userNames[id] = responseUser[j].first_name + " " + responseUser[j].last_name;
      userAvatars[id] = responseUser[j].photo_50;
    }
    
    // Print comments while removing duplicates (a possible problem on the other end)
    var prev = "";
    var curr = "";
    for (var j = 0; j < response.count; j++) {
      
      if (response.items[j]) {
        curr = response.items[j].from_id + response.items[j].date + response.items[j].text;
        
        if (curr !== prev) {
          var id = response.items[j].from_id;
          var commentDate = new Date(response.items[j].date * 1000);
          
          sheetComments.getRange(lineNumberOnSheet, 1).setValue(userNames[id]);
          sheetComments.getRange(lineNumberOnSheet, 2).setValue(commentDate);
          sheetComments.getRange(lineNumberOnSheet, 3).setValue(response.items[j].text);
          sheetComments.getRange(lineNumberOnSheet, 4).setValue(userAvatars[id]);
          
          lineNumberOnSheet++;
          prev = curr;
        }
      }
    }
    
    // Sleep until the script has run for a full second in total
    var timeElapsed = Date.now() - timeBegin;
    if (timeElapsed < 1000) {
      Utilities.sleep(1000 - timeElapsed);
    }
  }

  return [numberOfComments, lineNumberOnSheet];
}
