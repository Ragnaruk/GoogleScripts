function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  
  menu.addItem('Подготовить таблицу к работе.', 'initializeActiveSheet');
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
  sheetOptions.getRange('A2').setValue('URL трансляции');
  sheetOptions.getRange('A3').setValue('Остаток итераций');
  
  sheetOptions.getRange('B1').setNote('После запуска скрипта в правой части экрана появится ссылка на авторизацию и предоставление прав в ВК.' +
                                      ' При успешной авторизации произойдет перенаправление на страницу с адресом формата:' +
                                      ' https://oauth.vk.com/blank.html#access_token=XXXXX&expires_in=0&user_id=39199554&state=123456.' +
                                      ' Ссылку на эту страницу нужно вставить в это поле.');
  sheetOptions.getRange('B2').setNote('URL трансляции в формате: https://vk.com/videoXXXXX_XXXXX. Его можно получить под видео: Поделиться -> Экспортировать -> Прямая ссылка.');
  sheetOptions.getRange('B3').setNote('Количество оставшихся итераций выполнения программы. В это поле следует вписать натуральное число, ' +
                                      'и, если все остальные поля заполнены, скрипт начнет работу.');
  
  sheetOptions.getRange('A1:A3').setFontWeight('bold');
  sheetOptions.getRange('A1:B3').setBorder(true, true, true, true, true, true);
  
  sheetOptions.autoResizeColumn(1);
  
  getVkToken();
  
  ScriptApp.newTrigger('onSheetEdit')
  .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
  .onEdit()
  .create();
  
  var ui = SpreadsheetApp.getUi();
  ui.alert("Наведите курсор на ячейки B1:B3 для того, чтобы показать примечания.");
}

function onSheetEdit() {
  var cells = sheetOptions.getRange('B1:B3').getValues();
  
  var isReady = true;
  
  // Check whether all fields are filled
  for (var i = 0; i < cells.length && isReady; i++) {
    if (cells[i].toString().length === 0) {
      isReady = false;
    }
  }
  
  if (isReady) {
    // Parse value fields
    var userToken = sheetOptions.getRange('B1').getValue().toString();
    userToken = userToken.substring(userToken.search('#access_token=') + '#access_token='.length, userToken.search('&expires_in'));
    
    var videoUrl = sheetOptions.getRange('B2').getValue().toString();
    videoUrl = videoUrl.substring(videoUrl.search('https://vk.com/video') + 'https://vk.com/video'.length).split('_');
    var ownerId = videoUrl[0];
    var videoId = videoUrl[1];
    
    var offset = 0;
  
    // Create an empty comments sheet
    if (sheetComments != null) {
      sheetActive.deleteSheet(sheetComments);
    }
    
    sheetComments = sheetActive.insertSheet();
    sheetComments.setName('Комментарии');
    
    while (sheetOptions.getRange('B3').getValue() > 0) {
      sheetOptions.getRange('B3').setValue(sheetOptions.getRange('B3').getValue() - 1);
      offset = receiveVKComments(userToken, ownerId, videoId, offset);
      
      Utilities.sleep(1000);
    }
  }
}

function receiveVKComments(userToken, ownerId, videoId, offset) {
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
    
    for (var j = 0; j < response.count; j++) {
      if (response.items[j]) {
        var id = response.items[j].from_id
        
        sheetComments.getRange(1 + j + offset, 1).setValue(userNames[id]);
        sheetComments.getRange(1 + j + offset, 2).setValue(response.items[j].date);
        sheetComments.getRange(1 + j + offset, 3).setValue(response.items[j].text);
        sheetComments.getRange(1 + j + offset, 4).setValue(userAvatars[id]);
      }
    }
    
    var timeElapsed = Date.now() - timeBegin;
    if (timeElapsed < 1000) {
      Utilities.sleep(1100 - timeElapsed);
    }
  }

  return numberOfComments;
}
