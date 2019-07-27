function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  
  menu.addItem('Создать формы', 'createForms');
  menu.addToUi();
}

function createForms() {
  var currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetName = currentSpreadsheet.getName();
  var answerSheet = currentSpreadsheet.getSheets()[0];
  var studentSheet = currentSpreadsheet.getSheets()[1];
  
  // Get all student emails
  var studentEmails = [];
  for (var lineNumber = 1; studentSheet.getRange('A' + lineNumber).getValue() != ''; lineNumber++) {
    var studentEmail = studentSheet.getRange('A' + lineNumber).getValue();
    Logger.log('studentEmail: ' + studentEmail);
    
    studentEmails.push(studentEmail);
  }
  
  // Create a unique form for every student
  for (var i = 0; i < studentEmails.length; i++) {
    var formName = spreadsheetName + ' - ' + studentEmails[i];
    var form = FormApp.create(formName);
    Logger.log('formName: ' + formName);
    Logger.log('Published URL: ' + form.getPublishedUrl());
    Logger.log('Editor URL: ' + form.getEditUrl());
    
    // Set form options
    form.setCollectEmail(true);
    form.setLimitOneResponsePerUser(true);
    form.setRequireLogin(true);
    form.setDestination(FormApp.DestinationType.SPREADSHEET, currentSpreadsheet.getId());
    
    for (var lineNumber = 1; answerSheet.getRange('A' + lineNumber).getValue() != ''; lineNumber++) {
      var questionTitle = answerSheet.getRange('A' + lineNumber).getValue();
      var requiredNumberOfAnswersInQuestion = answerSheet.getRange('B' + lineNumber).getValue();
      var numbersOfEssentialQuestions = answerSheet.getRange('C' + lineNumber).getValue().toString().split(';');
      Logger.log('questionTitle: ' + questionTitle);
      Logger.log('requiredNumberOfAnswersInQuestion: ' + requiredNumberOfAnswersInQuestion);
      Logger.log('numbersOfEssentialQuestions: ' + numbersOfEssentialQuestions);
      
      // Get all answers to a question
      var allAnswersToQuestion = [];
      for (var j = 3; j < 26 && answerSheet.getRange(String.fromCharCode(65 + j) + lineNumber).getValue() != ''; j++) {
        var columnLetter = String.fromCharCode(65 + j);
        var answer = answerSheet.getRange(columnLetter + lineNumber).getValue();
        
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
      form.addMultipleChoiceItem()
      .setTitle(questionTitle)
      .setChoiceValues(requiredAnswersToQuestion)
      .showOtherOption(false);
    }
  }
}
