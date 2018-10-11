import base64
import json
import requests
import sys

templateName = sys.argv[1]
filename = sys.argv[2]
auth = (sys.argv[3], sys.argv[4])

headers = {'Content-Type':'application/json'}

bytes = []

f = open(filename, 'rb')
b = f.read(1)
while (b != ''):
  bytes.append(b)
  b = f.read(1)
print len(bytes)
bytearray = b''.join(bytes)
payload = {"data":{"templateName":templateName, "result":base64.b64encode(bytearray)}}
r = requests.post('http://localhost:5516/api/extension/templateImportExcel/import', json.dumps(payload), auth=auth, headers=headers)
print r
