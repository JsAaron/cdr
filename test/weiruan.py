
import json
import requests
import urllib
import uuid


# for microsoft translation
SUBKEY = '703c3c2419304a03b355e4c2cb7b964b'
ENDPOINT = 'https://api.cognitive.microsofttranslator.com'
DICTPATH = '/dictionary/lookup?api-version=3.0'
TRANSLATEPATH = '/translate?api-version=3.0'

wordlist = ["把图片3往左边移动10厘米,再往右移一点点,还要放大点"]
# wordlist = ["This is a table, I'm going to read, you go slow."]

params = '&from=zh&to=en'
# params = '&from=en&to=zh'
constructed_url = ENDPOINT + TRANSLATEPATH + params
headers = {
    'Ocp-Apim-Subscription-Key': SUBKEY,
    'Content-type': 'application/json',
    'X-ClientTraceId': str(uuid.uuid4())
}
body = []
for theword in wordlist:
    body.append({ 'text': theword })

request = requests.post(constructed_url, headers=headers, json=body)
response = request.json()


jsonstr = json.dumps(response, sort_keys=True, indent=4,
              ensure_ascii=False, separators=(',', ': '))
results = json.loads(jsonstr)

ret = []
if len(results) > 0 :
    for index, result in enumerate(results):
        transtext = wordlist[index]
        items = result['translations']
        if items != None:
            transtext = items[0]['text']
        ret.append(transtext)

  
print(ret)