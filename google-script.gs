function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  
  menu.addItem('Подготовить таблицу к работе', 'initializeActiveSheet');
  menu.addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showVKAuthSidebar() {
  var authorizationUrl = 'https://oauth.vk.com/authorize?client_id=6947304&display=page&' +
    'redirect_uri=https://oauth.vk.com/blank.html&scope=video,offline&response_type=token&v=5.95&state=123456';
  
  var template = HtmlService.createTemplate(
    '<style type="text/css">' +
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
    'После нажатия на кнопку ниже браузер перенаправит вас на сайт ВК,' +
    ' где нужно будет авторизоваться и разрешить приложению доступ к вашим данным.' +
    '</p>' +
    '<p>' +
    'После окончания авторизации вас перенаправит на страницу с надписью:' +
    ' "Пожалуйста, не копируйте данные из адресной строки для сторонних сайтов.' +
    ' Таким образом Вы можете потерять доступ к Вашему аккаунту."' +
    '</p>' +
    '<p>' +
    'Вам нужно будет скопировать адрес этой страницы из адресной строки в ячейку B1.' +
    '</p>' +
    '<button id="oauth" onclick="redirect()">Авторизоваться в ВК</button>'
  );
  
  template.authorizationUrl = authorizationUrl;
  var page = template.evaluate();
  
  SpreadsheetApp.getUi().showSidebar(page);
}

var sheetActive = SpreadsheetApp.getActiveSpreadsheet();
var sheetOptions = sheetActive.getSheets()[0];
var sheetComments = sheetActive.getSheetByName('Комментарии');

function initializeActiveSheet() {
  // Delete all triggers and create an onEdit trigger that will launch the main logic
  deleteAllTriggers();
  
  // Initialize options sheet
  sheetOptions = sheetActive.getSheets()[0];
  sheetOptions.setName('Настройки');
  
  sheetActive.setActiveSheet(sheetOptions);
  
  sheetOptions.getRange('A1').setValue('Ссылка авторизации');
  sheetOptions.getRange('A2').setValue('Прямая ссылка на трансляцию');
  sheetOptions.getRange('A3').setValue('Дата и время окончания');
  
  sheetOptions.getRange('B3').setValue(new Date());
  sheetOptions.getRange('B3').setNumberFormat("dd.MM.yyyy hh:mm:ss");
  
  sheetOptions.getRange('B1').setNote(
    'После запуска скрипта в правой части экрана появится ссылка на авторизацию и предоставление прав в ВК.' +
    ' При успешной авторизации произойдет перенаправление на страницу с адресом формата:' +
    ' https://oauth.vk.com/blank.html#access_token=XXXXX&expires_in=0&user_id=39199554&state=123456.' +
    ' Ссылку на эту страницу нужно вставить в это поле.'
  );
  sheetOptions.getRange('B2').setNote(
    'URL трансляции в формате: https://vk.com/videoXXXXX_XXXXX.' +
    ' Его можно получить под видео: Поделиться -> Экспортировать -> Прямая ссылка.'
  );
  sheetOptions.getRange('B3').setNote(
    'Скрипт будет выполняться до тех пор, пока не наступит время, указанное в ячейке.'
  );
  
  sheetOptions.getRange('A1:A3').setFontWeight('bold');
  sheetOptions.getRange('A1:B3').setBorder(true, true, true, true, true, true);
  
  sheetOptions.autoResizeColumns(1, 2);
  
  
  showVKAuthSidebar();
  createOnEditTrigger();
}

function stopAllRunningScripts() {
  // Delete all comments triggers
  if (PropertiesService.getScriptProperties().getProperty('commentsTriggers')) {
    var commentsTriggers = JSON.parse(PropertiesService.getScriptProperties().getProperty('commentsTriggers'));
    
    for (var i = 0; i < commentsTriggers.length; i++) {
      var triggerUid = commentsTriggers[i];
      
      ScriptApp.getProjectTriggers().some(function (trigger) {
        if (trigger.getUniqueId() === triggerUid) {
          ScriptApp.deleteTrigger(trigger);
        }
      })
    }
  }
  
  // Delete all global properties
  PropertiesService.getScriptProperties().deleteAllProperties();
}

function deleteAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

function createOnEditTrigger() {
  ScriptApp.newTrigger('onSheetEdit')
  .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
  .onEdit()
  .create();
}

// Launches every time sheet is edited
function onSheetEdit() {
  // Stop all currently running scripts and remove their triggers
  stopAllRunningScripts();
  
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
    // Parse value fields
    var userToken = sheetOptions.getRange('B1').getValue().toString();
    userToken = userToken.substring(userToken.search('#access_token=') + '#access_token='.length, userToken.search('&expires_in'));
    
    var videoUrl = sheetOptions.getRange('B2').getValue().toString();
    videoUrl = videoUrl.substring(videoUrl.search('https://vk.com/video') + 'https://vk.com/video'.length).split('_');
    var ownerId = videoUrl[0];
    var videoId = videoUrl[1];
    
    // Create variables for offset (number of the first comment to receive) and line number on the comments sheet
    var offset = 0;
    var lineNumberOnSheet = 2;
    
    // Write all script variables to a global property
    var properties = [];
    properties.push(userToken);
    properties.push(ownerId);
    properties.push(videoId);
    properties.push(offset);
    properties.push(lineNumberOnSheet);
    
    // Create an id of the script from a timestamp
    var scriptId = (new Date).valueOf();
    
    PropertiesService.getScriptProperties().setProperty(scriptId, JSON.stringify(properties));
    
    // Create an empty comments sheet and prepare it for writin
    sheetComments = sheetActive.getSheetByName('Комментарии');
    
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
    
    // Create a trigger to launch scripts every 5 minutes to avoid 6 minute time limit of free accounts
    createTimeTriggerToReceiveComments(scriptId)
  }
}

function createTimeTriggerToReceiveComments(scriptId) {
  // Run function once while time-based trigger is being created creating
  startReceiveCommentsLoop(scriptId);
  
  // Create a trigger that will start in 5 minutes and run every 5 minute since then
  var trigger = ScriptApp.newTrigger('startReceiveCommentsLoop')
  .timeBased()
  .everyMinutes(5)
  .create();
  
  // Create a global variable to pass scriptId to triggered function
  var triggerUid = trigger.getUniqueId();
  
  PropertiesService.getScriptProperties().setProperty(triggerUid, scriptId.toString());
  
  // Update a global variable to contain trigger uid
  if (PropertiesService.getScriptProperties().getProperty('commentsTriggers')) {
    var commentsTriggers = JSON.parse(PropertiesService.getScriptProperties().getProperty('commentsTriggers'));
  } else {
    var commentsTriggers = [];
  }
  
  commentsTriggers.push(triggerUid);
  
  PropertiesService.getScriptProperties().setProperty('commentsTriggers', JSON.stringify(commentsTriggers));
}

// Create a loop to launch main logic
function startReceiveCommentsLoop(e) {
  // If this function is launch from trigger, get scriptId from it else get it directly
  if (e.triggerUid) {
    var triggerUid = e.triggerUid;
    var scriptId = PropertiesService.getScriptProperties().getProperty(triggerUid);
  } else {
    var scriptId = e;
  }
  
  if (PropertiesService.getScriptProperties().getProperty(scriptId)) {
    var properties = getGlobalPropertiesOfAScript(scriptId);
    var timeBegin = Date.now();
    
    // While script has run less than 5 minutes, current date is less than date in the cell
    // and while global properties of the script exist
    while (
      Date.now() - timeBegin < 1000 * 60 * 5
      && sheetOptions.getRange('B3').getValue().valueOf() > Date.now()
      && PropertiesService.getScriptProperties().getProperty(scriptId)
    ) {
      // Pass: userToken, ownerId, videoId, offset, lineNumberOnSheet
      // Receive: offset, lineNumberOnSheet
      var resp = receiveComments(properties[0], properties[1], properties[2], properties[3], properties[4]);
      properties[3] = resp[0];
      properties[4] = resp[1];
      
      // Update global properties of the script
      PropertiesService.getScriptProperties().setProperty(scriptId, JSON.stringify(properties));
      
      Utilities.sleep(1000);
    }
  }
}

function getGlobalPropertiesOfAScript(scriptId) {
  return JSON.parse(PropertiesService.getScriptProperties().getProperty(scriptId));
}

// Main logic function
function receiveComments(userToken, ownerId, videoId, offset, lineNumberOnSheet) {
  var url = 'https://api.vk.com/method/video.getComments?count=100&sort=asc&' +
    'owner_id=' + ownerId + '&video_id=' + videoId + '&offset=' + 0 + '&access_token=' + userToken + '&v=5.95';
  
  var response = JSON.parse(UrlFetchApp.fetch(url).getContentText()).response;
  
  var numberOfComments = response.count;
  
  if (numberOfComments - offset > 50) {
    offset = numberOfComments - 50;
  }
  
  for (var i = offset; i < numberOfComments; i += 100) {
    // Time the speed of completion to avoid the limit of 3 requests / second
    var timeBegin = Date.now()
    
    if (numberOfComments - i < 100) {
      var count = numberOfComments - i;
    } else {
      var count = 100;
    }
    
    var url = 'https://api.vk.com/method/video.getComments?count=100&sort=asc&' +
      'owner_id=' + ownerId + '&video_id=' + videoId + '&offset=' + i + '&access_token=' + userToken + '&v=5.95';
    var response = JSON.parse(UrlFetchApp.fetch(url).getContentText()).response;

    // Add all user ids to a string and separate them by commas
    var user_ids = "";
    for (var j = 0; j < response.count; j++) {
      if (response.items[j]) {
        user_ids += response.items[j].from_id + ",";
      }
    }
    user_ids = user_ids.slice(0, -1);
    
    var urlUser = 'https://api.vk.com/method/users.get?' +
      'user_ids=' + user_ids + '&fields=photo_50&access_token=' + userToken + '&v=5.95';
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
      // If both response and comments sheet exist
      if (response.items[j] && sheetComments) {
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
