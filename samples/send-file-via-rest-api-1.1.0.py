import base64
import json
import requests
import sys

targetFolderId = sys.argv[1]
templateName = sys.argv[2]
filename = sys.argv[3]
auth = (sys.argv[4], sys.argv[5])

headers = {'Content-Type':'application/json'}

bytes = []

f = open(filename, 'rb')
b = f.read(1)
while (b != ''):
  bytes.append(b)
  b = f.read(1)
print len(bytes)
bytearray = b''.join(bytes)
payload = {"data":{"targetFolderId":targetFolderId, "templateName":templateName, "result":base64.b64encode(bytearray)}}
r = requests.post('http://localhost:5516/api/extension/templateImportExcel/import', json.dumps(payload), auth=auth, headers=headers)
print r
