# -*- coding: utf-8 -*-
# python slack.py username channel test file
# ex.
# python slack.py andes #w0 test

import sys
import slackweb
import codecs

#url= "https://hooks.slack.com/services/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"
#url= "https://hooks.slack.com/services/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"
url= "https://hooks.slack.com/services/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"
slack = slackweb.Slack(url = url)

args = sys.argv
#title = "Sushi"
print(args[0])
username = args[1]
print("user:", username)
channel = args[2]
print("channel:", channel)
text = args[3]
print("text:", text)
file = args[4]
print("file:", file)
#read_text = open(file).read()
read_text = ''
for line in open(file).readlines():
    read_text += line.rstrip() + "\n"

print("len(read_text)=", len(read_text))
if read_text != "":
    if len(read_text) > 4000:
        print("--- cut ---")
        text += "\n```"
        text += read_text[:1800]
        text += "```\n--- cut:len({0}) ---\n```".format(len(read_text))
        text += read_text[-1800:]+ "\n```"
    else:
        text += "\n```" + read_text + "\n```"
print("slack.notify({0}, {1}, {2})".format(args[3], username, channel), end="." , flush=None)
slack.notify(text = text, username = username, channel = channel)
print("ok")

#attachments = []
#attachment = {"text": "Eating *right now!*"}
#attachments.append(attachment)
#slack.notify(attachments=attachments)
