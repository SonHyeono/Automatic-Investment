import requests


def post_message(channel, text):
    response = requests.post("https://slack.com/api/chat.postMessage",
                             headers={"Authorization": "Bearer " +
                                      "***********************"},
                             data={"channel": channel, "text": text}
                             )
    print(response)


def post_token():
    return "****************************"
