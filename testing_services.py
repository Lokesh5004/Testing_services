import requests
from datetime import datetime
import win32com.client

email_config = {
    'to': 'lokesh.vudugu.ext@nokia.com',
    'cc': 'lokesh.vudugu.ext@nokia.com',
}

outlook = win32com.client.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
urls = ['http://fiestpm012.nsn-intra.net:8080/qcbin',
        'http://fiestpm012.nsn-intra.net:8080/qcbin/SiteAdmin.jsp',
        'http://fiestpm012.nsn-intra.net:8080/qcbin/start_a.jsp?projectType=LabManagement',
        'http://fiestpm009.nsn-intra.net/LoadTest',
        'http://fihe3nok1145.nsn-intra.net/LoadTest',
        'http://esptv011.nsn-intra.net:8080/SiteScope',
        'https://fihe3nok1183.nsn-intra.net:5814/autopass1c',
        'https://demutpm009.nsn-intra.net:5814/autopass/login_input1']

data = {}
for url in urls:
    print("Testing URL:\t" + url)
    r = requests.get(url, verify=False)
    data[url] = 'up' if r.status_code  == 200 else 'down'

body = 'Hi All, <br> '
if 'down' in data.values():
    body += '<p style="background-color:##FFF000">The below mentioned URL('+", ".join([url for url, status in data.items() if status == "down"])+') is down please take required action </p><br><br>'
body += 'All URL''s are up and running fine.'
body += '<br><br> <table border="1" cellspacing="0" cellpadding="3">'
body += '<tr style="background-color:#009BFF"><th colspan="2">Testing Services (Date- '+datetime.today().strftime('%d.%m.%Y')+' Time- '+ datetime.today().strftime("%I:%M %p") +'</th></tr>'
body += '<tr style="background-color:#72B6E2"><th>URL</th><th>Status</th></tr>'
for url, status in data.items():
    body += '<tr><td><a href="'+url+'">'+ url +'</a></td><td>'
    body += '<p style="color:' + ('#11DC00' if status == 'up' else '#DC0000') + '"><b>'+status+'</b></p>'
    body += '</td></tr>'
body += '</table>'

mail.To = email_config['to']
mail.CC = email_config['cc']
mail.Subject = '!!! '+ ('GREEN' if len(set(data.values())) == 1 and next(iter(set(data.values()))) == 'up' else 'RED' ) +' ALERT!!! Testing Services URL Monitoring report '+ datetime.today().strftime('%d-%m-%Y %I:%M %p')
mail.HTMLBody = body
mail.Send()