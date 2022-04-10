import jwt
import requests
import json
from time import time

API_KEY = 'uZlAleG7QNWNtebfcAAKKQ'
API_SEC = 'hdc1bkE84b2WnrIIuyt5C8npKdxQErv9jGd8'

# your zoom live meeting id, it is optional though
#meetingId = 83781439159

#userId = 'you can get your user Id by running the getusers()'
userId = 'divyanshu@intellipaat.com'
# create a function to generate a token using the pyjwt library
def generateToken():
    token = jwt.encode(
        # Create a payload of the token containing API Key & expiration time
        {'iss': API_KEY, 'exp': time() + 5000},
        # Secret used to generate token signature
        API_SEC,
        # Specify the hashing alg
        algorithm='HS256'
        # Convert token to utf-8
    )
    return token
    # send a request with headers including a token


#fetching zoom meeting participants of the live meeting

# def getMeetingParticipants():
#     headers = {'authorization': 'Bearer %s' % generateToken(),
#                'content-type': 'application/json'}
#     r = requests.get(
#         f'https://api.zoom.us/v2/metrics/meetings/{meetingId}/participants', headers=headers)
#     print("\n fetching zoom meeting participants of the live meeting ... \n")

#     # you need zoom premium subscription to get this detail, also it might not work as i haven't checked yet(coz i don't have zoom premium account)

#     print(r.text)


# # this is the json data that you need to fill as per your requirement to create zoom meeting, look up here for documentation
# # https://marketplace.zoom.us/docs/api-reference/zoom-api/meetings/meetingcreate


meetingdetails = {"topic": input("Topic: "),
                  "type": 2,
                  "start_time": input("YYYY-MM-DD: ")+"T"+input("HH:MM:SS: "),
                  "duration": input("Number of Minutes: "),
                  "timezone": "India",
                  "agenda": "test",

                  "recurrence": {"type": 1,
                                 "repeat_interval": 1
                                 },
                  "settings": {"host_video": "true",
                               "participant_video": "true",
                               "join_before_host": "False",
                               "mute_upon_entry": "False",
                               "watermark": "true",
                               "audio": "voip",
                               "auto_recording": "cloud"
                               }
                  }


def createMeeting():
    headers = {'authorization': 'Bearer %s' % generateToken(),
               'content-type': 'application/json'}
    r = requests.post(
        f'https://api.zoom.us/v2/users/{userId}/meetings', headers=headers, data=json.dumps(meetingdetails))

    print("\n creating zoom meeting ... \n")
    #print(r.text)
    return r

value = createMeeting()

v = value.text
v = v[1:-1]
v = v.split(',')
d11={}
for i in v:
    x=i.split(':',1)
    x[0],x[1] = x[0][1:-1],x[1][1:-1]
    d11[x[0]]=x[1]
    
print(d11['join_url'])

import pandas as pd
emails = pd.read_excel(r'NameEmails.xlsx')
bcc = list(emails['E mail'])

import smtplib, ssl
import getpass

port = 465  # For SSL
smtp_server = "smtp.gmail.com"
sender_email = input("Enter Sender Email")  # Enter your address
receiver_email = input("Enter Receiver Email")  # Enter receiver address
password = getpass.getpass('Password:')
sub = input("Enter Subject line")
message_text = "Just a test: " + d11['join_url']
message = "From: %s\r\n" % sender_email + "To: %s\r\n" % receiver_email + "Subject: %s\r\n" % sub + "\r\n" + message_text
total = [sender_email] + bcc

context = ssl.create_default_context()
with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
    server.login(sender_email, password)
    server.sendmail(sender_email, total, message)