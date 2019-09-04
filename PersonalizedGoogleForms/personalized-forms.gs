function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  
  menu.addItem('Инициализировать таблицу', 'initializeActiveSheet');
  menu.addItem('Создать формы', 'createForms');
  menu.addToUi();
}

function initializeActiveSheet() {
  var currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var questionSheet = currentSpreadsheet.getSheetByName('Вопросы') ? currentSpreadsheet.getSheetByName('Вопросы') : currentSpreadsheet.getSheets()[0];
  var studentSheet = currentSpreadsheet.getSheetByName('Студенты') ? currentSpreadsheet.getSheetByName('Студенты') : currentSpreadsheet.getSheets()[1];
  
  questionSheet.setName('Вопросы');
  
  questionSheet.getRange('A1').setValue('Текст вопроса');
  questionSheet.getRange('B1').setValue('Количество необходимых вариантов ответов');
  questionSheet.getRange('C1').setValue('Номера обязательных вариантов ответов через \';\'');
  questionSheet.getRange('D1').setValue('Варианты ответов');
  
  questionSheet.getRange('A1:D1').setFontWeight('bold');
  questionSheet.autoResizeColumns(1, 4);
  
  studentSheet.setName('Студенты');
  
  studentSheet.getRange('A1').setValue('Email адрес студента');
  studentSheet.getRange('B1').setValue('Ссылка на персональную форму');
  
  studentSheet.getRange('A1:B1').setFontWeight('bold');
  studentSheet.autoResizeColumns(1, 2);
  
  PropertiesService.getScriptProperties().setProperty('isInitialized', true);
}

function createForms() {
  if (!PropertiesService.getScriptProperties().getProperty('isInitialized')) {
    SpreadsheetApp.getUi().alert('Пожалуйста, инициализируйте таблицу перед началом работы.');
    return;
  }
  
  var currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetName = currentSpreadsheet.getName();
  var questionSheet = currentSpreadsheet.getSheetByName('Вопросы') ? currentSpreadsheet.getSheetByName('Вопросы') : currentSpreadsheet.getSheets()[0];
  var studentSheet = currentSpreadsheet.getSheetByName('Студенты') ? currentSpreadsheet.getSheetByName('Студенты') : currentSpreadsheet.getSheets()[1];
  
  // Get all student emails
  var studentEmails = [];
  for (var lineNumber = 1; studentSheet.getRange('A' + lineNumber).getValue() != ''; lineNumber++) {
    var studentEmail = studentSheet.getRange('A' + lineNumber).getValue();
    Logger.log('studentEmail: ' + studentEmail);
    
    studentEmails.push(studentEmail);
  }
  
  // Create a unique form for every student
  for (var i = 1; i < studentEmails.length; i++) {
    var formName = spreadsheetName + ' - ' + studentEmails[i];
    var form = FormApp.create(formName);
    Logger.log('formName: ' + formName);
    Logger.log('formId: ' + form.getId());
    Logger.log('Published URL: ' + form.getPublishedUrl());
    Logger.log('Editor URL: ' + form.getEditUrl());
    
//    // Send an email to the student
//    informStudent(studentEmails[i], form.getPublishedUrl());
    
    // Print form url to the sheet
    studentSheet.getRange('B' + (i + 1)).setValue(form.getPublishedUrl());
    
    // Add onSumbit trigger to delete the form
    addTriggerToForm(form.getId());
    
    // Set form options
    form.setCollectEmail(true);
    form.setLimitOneResponsePerUser(true);
    form.setRequireLogin(true);
    form.setDestination(FormApp.DestinationType.SPREADSHEET, currentSpreadsheet.getId());
    
    // Populate form
    for (var lineNumber = 2; questionSheet.getRange('A' + lineNumber).getValue() != ''; lineNumber++) {
      var questionTitle = questionSheet.getRange('A' + lineNumber).getValue();
      var requiredNumberOfAnswersInQuestion = questionSheet.getRange('B' + lineNumber).getValue();
      var numbersOfEssentialQuestions = questionSheet.getRange('C' + lineNumber).getValue().toString().split(';');
      Logger.log('questionTitle: ' + questionTitle);
      Logger.log('requiredNumberOfAnswersInQuestion: ' + requiredNumberOfAnswersInQuestion);
      Logger.log('numbersOfEssentialQuestions: ' + numbersOfEssentialQuestions);
      
      // Get all answers to a question
      var allAnswersToQuestion = [];
      for (var j = 3; j < 26 && questionSheet.getRange(String.fromCharCode(65 + j) + lineNumber).getValue() != ''; j++) {
        var columnLetter = String.fromCharCode(65 + j);
        var answer = questionSheet.getRange(columnLetter + lineNumber).getValue();
        
        allAnswersToQuestion.push(answer);
      }
      Logger.log('allAnswersToQuestion: ' + allAnswersToQuestion);
      
      // Get required number of answers to a question
      var requiredAnswersToQuestion = [];
      if (allAnswersToQuestion.length <= requiredNumberOfAnswersInQuestion) {
        requiredAnswersToQuestion = allAnswersToQuestion;
      } else {
        // Sort array in descending order
        numbersOfEssentialQuestions.sort(function (a, b) { return b - a; });
        
        // Get required answers
        for (var j = 0; j < numbersOfEssentialQuestions.length; j++) {
          var essentialQuestion = allAnswersToQuestion[numbersOfEssentialQuestions[j] - 1];
          Logger.log('essentialQuestion: ' + essentialQuestion);
          
          requiredAnswersToQuestion.push(essentialQuestion);
          
          // Remove used question from array
          allAnswersToQuestion.splice(numbersOfEssentialQuestions[j] - 1, 1);
        }
        
        // Get remaining answers
        while (requiredAnswersToQuestion.length < requiredNumberOfAnswersInQuestion) {
          var randomQuestionNumber = Math.floor(Math.random() * allAnswersToQuestion.length);
          
          requiredAnswersToQuestion.push(allAnswersToQuestion[randomQuestionNumber]);
          allAnswersToQuestion.splice(randomQuestionNumber, 1);
        }
      }
      
      // Shuffle answers
      for (var j = requiredAnswersToQuestion.length - 1; j > 0; j--) {
        var k = Math.floor(Math.random() * (j + 1));
        [requiredAnswersToQuestion[j], requiredAnswersToQuestion[k]] = [requiredAnswersToQuestion[k], requiredAnswersToQuestion[j]];
      }
      
      // Add a question to the form
      form.addCheckboxItem()
      .setTitle(questionTitle)
      .setChoiceValues(requiredAnswersToQuestion)
      .showOtherOption(false);
    }
  }
}

function informStudent(email, formURL) {
  MailApp.sendEmail(email, 'Новая форма', 'Вам был открыт доступ к новой форме: ' + formURL);
}

function addTriggerToForm(formId) {
  var trigger = ScriptApp.newTrigger('onFormSubmit')
  .forForm(formId)
  .onFormSubmit()
  .create();
  
  Logger.log('Created triggerId: ' + trigger.getUniqueId());
  
  PropertiesService.getScriptProperties().setProperty(formId, trigger.getUniqueId());
}

function deleteTriggerById(formId) {
  var triggerId = PropertiesService.getScriptProperties().getProperty(formId);
  var triggers = ScriptApp.getProjectTriggers();
  
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getUniqueId() === triggerId) {
      ScriptApp.deleteTrigger(triggers[i]);
      Logger.log('Deleted triggerId: ' + triggerId);
    }
  }
}

function onFormSubmit(e) {
  var form = e.source;
  
  form.setAcceptingResponses(false);
  
  deleteTriggerById(form.getId());
}
