var RabbitMQHost = "";
var RabbitMQPort = "";
var RabbitMQCredentials = "";
var RabbitMQQueueName = "";
var RabbitMQCallbackQueueName = "";

function main() {
    var message = JSON.stringify({
        "xqueue_header": {
            "submission_id": 72,
            "submission_key": "ffcd933556c926a307c45e0af5131995"
        },
        "xqueue_body": {
            "student_info": {
                "anonymous_student_id": "106ecd878f4148a5cabb6bbb0979b730",
                "submission_time": (new Date).toString(),
                "random_seed": Math.floor(Math.random() * 10000)
            },
            "student_response": "5",
            "grader_payload": "2"
        }
    });
    
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
        \'payload\': \'' + message + '\', \
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
        \'count\': 1, \
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
    
    Logger.log(result);
}

