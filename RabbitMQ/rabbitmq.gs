var RabbitMQHost = "";
var RabbitMQPort = "";
var RabbitMQCredentials = "";
var RabbitMQQueueName = "";
var RabbitMQCallbackQueueName = "";

var submissionID = 0;

function main() {
    gradeAnswer("test", 10)
}

function createTrigger() {
    var trigger = ScriptApp.newTrigger('receiveMessageFromRabbitMQ')
        .timeBased()
        .everyMinutes(1)
        .create();
}

function gradeAnswer(answer, studentID) {
    if (!PropertiesService.getScriptProperties().getProperty("submissionID")) {
        PropertiesService.getScriptProperties().setProperty("submissionID", 0);
    }
    var submissionID = PropertiesService.getScriptProperties().getProperty("submissionID");
    
    var message = {
        "xqueue_header": {
            "submission_id": submissionID,
            "submission_key": (new Date().getTime()).toString()
        },
        "xqueue_body": {
            "student_response": answer.toString(),
            "grader_payload": "2"
        }
    };
    
    PropertiesService.getScriptProperties().setProperty("submissionID", ++submissionID);
    PropertiesService.getScriptProperties().setProperty(JSON.stringify(message.xqueue_header), studentID.toString());
    Logger.log("message: " + JSON.stringify(message));
    
    sendMessageToRabbitMQ(message);
}

function sendMessageToRabbitMQ(message) {
    var payload = '{ \
    \'vhost\': \'/\', \
    \'name\': \'amq.default\', \
    \'properties\': { \
      \'delivery_mode\': 1, \
      \'headers\': {}, \
      \'reply_to\': \'' + RabbitMQCallbackQueueName + '\' \
    }, \
    \'routing_key\': \'' + RabbitMQQueueName + '\', \
    \'delivery_mode\': \'1\', \
    \'payload\': \'' + JSON.stringify(message) + '\', \
    \'headers\': {}, \
    \'props\': {}, \
    \'payload_encoding\': \'string\' \
  }';
    
    var options = {
        "method": "POST",
        "payload": payload,
        "headers": {
            "Authorization": "Basic " + Utilities.base64Encode(RabbitMQCredentials),
            "Content-Type": "application/json;charset=UTF-8"
        },
        "muteHttpExceptions": false
    };
    
    var result = UrlFetchApp.fetch("http://" + RabbitMQHost + ":" + RabbitMQPort + "/api/exchanges/%2F/amq.default/publish", options);
    
    Logger.log(result);
}

function receiveMessageFromRabbitMQ() {
    var payload = '{ \
    \'count\': 10, \
    \'ackmode\': \'ack_requeue_false\', \
    \'encoding\': \'auto\', \
    \'truncate\': 5000 \
  }';
    
    var options = {
        "method": "POST",
        "payload": payload,
        "headers": {
            "Authorization": "Basic " + Utilities.base64Encode(RabbitMQCredentials),
            "Content-Type": "application/json;charset=UTF-8"
        },
        "muteHttpExceptions": false
    };
    
    var result = UrlFetchApp.fetch("http://" + RabbitMQHost + ":" + RabbitMQPort + "/api/queues/%2f/" + RabbitMQCallbackQueueName + "/get", options);
    
    if (result.getContentText().length > 2) {
        Logger.log("Получено сообщение: " + result);
        var messages = JSON.parse(result.getContentText());
        
        for (var i = 0; i < messages.length; i++) {
            payload = messages[i].payload.replace("\'", "\"", "g").replace("True", "true", "g").replace("False", "false", "g");
            payload = JSON.parse(payload);
            if (PropertiesService.getScriptProperties().getProperty(JSON.stringify(payload.xqueue_header))) {
                var studentID = PropertiesService.getScriptProperties().getProperty(JSON.stringify(payload.xqueue_header));
                
                processGrade(payload.xqueue_body, studentID)
            } else {
                Logger.log("Получено неопознанное сообщение: " + payload);
            }
        }
    } else {
        Logger.log("Сообщений в очереди нет.")
    }
}

function processGrade(grade, studentID) {
    Logger.log("grade: " + JSON.stringify(grade));
    Logger.log("studentID: " + studentID);
}
