var formId = '';
var sourceFileId = '';
var destinationFolderId = '';

// Add a trigger which listens for form submit
function addTrigger() {
    ScriptApp.newTrigger('onSubmit')
        .forForm(formId)
        .onFormSubmit()
        .create();
}

// On form submit copy a file from Google Drive and share a copy with a person who left their email in the form
function onSubmit(e) {
    var email = e.response.getRespondentEmail();
    var newFileName = 'Some Name — ' + (new Date()).valueOf();
    var destinationFolder = DriveApp.getFolderById(destinationFolderId);
    
    var file = DriveApp.getFileById(sourceFileId);
    var newFile = file.makeCopy(newFileName, destinationFolder);
    
    newFile.addEditor(email);
}
