#
# Copyright 2018 XEBIALABS
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#

from java.io import ByteArrayInputStream
from org.apache.commons.codec.binary import Base64
from org.apache.poi.xssf.usermodel import XSSFWorkbook

from templateImportExcel.TemplateImportExcelClientUtil import Template_Import_Excel_Client_Util

logger.info("Executing templateImportExcel/import-rest.py")

targetFolderId = request.entity['data']['targetFolderId']
templateName = request.entity['data']['templateName']
workbook = XSSFWorkbook(ByteArrayInputStream(Base64.decodeBase64(request.entity['data']['result'])))
print "created workbook\n"
client = Template_Import_Excel_Client_Util.create_client(workbook, targetFolderId, templateName, templateApi, phaseApi, taskApi)
client.convertWorkbookToTemplate()

logger.info("Exiting import-rest.py")
