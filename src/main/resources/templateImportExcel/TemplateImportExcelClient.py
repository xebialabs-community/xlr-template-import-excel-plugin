#
# Copyright 2018 XEBIALABS
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#

from java.util import Date

from com.xebialabs.xlrelease.domain import Phase
from com.xebialabs.xlrelease.domain import Release
from com.xebialabs.xlrelease.domain import Task
from com.xebialabs.xlrelease.domain.status import ReleaseStatus

import re

class Template_Import_Excel_Client(object):

    columnXref = {}
    currentPhaseTitle = None
    currentPhase = None
    currentTask = None
    template = None

    workbook = None
    templateName = None
    templateApi = None
    phaseApi = None
    taskApi = None

    durationMatcher = None

    def __init__(self, workbook, targetFolderId, templateName, templateApi, phaseApi, taskApi):
        self.columnXref = {}
        self.currentPhaseTitle = None
        self.currentPhase = None
        self.targetFolderId = targetFolderId
        self.templateName = templateName
        self.workbook = workbook
        self.templateApi = templateApi
        self.phaseApi = phaseApi
        self.taskApi = taskApi
        self.durationMatcher = re.compile("(?:([0-9]+)d)? *(?:([0-9]+)h)? *(?:([0-9]+)m)?")

    @staticmethod
    def create_client(workbook, targetFolderId, templateName, templateApi, phaseApi, taskApi):
        return Template_Import_Excel_Client(workbook, targetFolderId, templateName, templateApi, phaseApi, taskApi)

    def doHeaderRow(self, row):
        for cell in row:
            if cell.getRichStringCellValue().getString() == "Phase":
                print "Phase in column %d\n" % cell.getColumnIndex()
                self.columnXref['phase'] = cell.getColumnIndex()
            if cell.getRichStringCellValue().getString() == "Task":
                print "Task in column %d\n" % cell.getColumnIndex()
                self.columnXref['taskname'] = cell.getColumnIndex()
            if cell.getRichStringCellValue().getString() == "Type":
                print "Type in column %d\n" % cell.getColumnIndex()
                self.columnXref['tasktype'] = cell.getColumnIndex()
            if cell.getRichStringCellValue().getString() == "User":
                print "User in column %d\n" % cell.getColumnIndex()
                self.columnXref['user'] = cell.getColumnIndex()
            if cell.getRichStringCellValue().getString() == "Team":
                print "Team in column %d\n" % cell.getColumnIndex()
                self.columnXref['team'] = cell.getColumnIndex()
            if cell.getRichStringCellValue().getString() == "Duration":
                print "Duration in column %d\n" % cell.getColumnIndex()
                self.columnXref['duration'] = cell.getColumnIndex()
            if cell.getRichStringCellValue().getString() == "Description":
                print "Description in column %d\n" % cell.getColumnIndex()
                self.columnXref['description'] = cell.getColumnIndex()

    def doPhaseCell(self, cell):
        if cell:
            phaseTitle = cell.getRichStringCellValue().getString()
            if not self.currentPhaseTitle or phaseTitle != self.currentPhaseTitle:
                print "Adding a new phase\n"
                self.currentPhase = self.phaseApi.addPhase(self.template.getId(), self.phaseApi.newPhase(phaseTitle))

    def doTaskNameCell(self, cell):
        if cell:
            task = self.taskApi.newTask("xlrelease.Task")
            task.title = cell.getRichStringCellValue().getString()
            self.currentTask = task

    def doTaskTypeCell(self, task, cell):
        if cell:
            self.taskApi.changeTaskType(self.currentTask.getId(), cell.getRichStringCellValue().getString())

    def doUserCell(self, task, cell):
        if cell:
            self.taskApi.assignTask(self.currentTask.getId(), cell.getRichStringCellValue().getString())

    def doTeamCell(self, task, cell):
        if cell:
            print "To-do:  Add team to task\n"

    def doDescriptionCell(self, task, cell):
        if cell:
            task.description = cell.getRichStringCellValue().getString()

    def doDurationCell(self, task, cell):
        if cell:
            (days, hours, minutes) = self.durationMatcher.match(cell.getRichStringCellValue().getString()).groups()
            if days or hours or minutes:
                plannedDuration = 0
                if days:
                    plannedDuration += 60 * 60 * 24 * int(days)
                if hours:
                    plannedDuration += 60 * 60 * int(hours)
                if minutes:
                    plannedDuration += 60 * int(minutes)
                task.plannedDuration = plannedDuration

    def doTask(self, task):
        self.taskApi.addTask(self.currentPhase.getId(), task)
        print "Task %s added to phase\n" % task.title

    def doDataRow(self, row):
        self.currentTask = None
        if 'phase' in self.columnXref:
            self.doPhaseCell(row.getCell(self.columnXref['phase']))
        if 'taskname' in self.columnXref:
            self.doTaskNameCell(row.getCell(self.columnXref['taskname']))
        if 'description' in self.columnXref:
            self.doDescriptionCell(self.currentTask, row.getCell(self.columnXref['description']))
        if 'duration' in self.columnXref:
            self.doDurationCell(self.currentTask, row.getCell(self.columnXref['duration']))
        if 'team' in self.columnXref:
            self.doTeamCell(self.currentTask, row.getCell(self.columnXref['team']))
        if self.currentTask:
            self.doTask(self.currentTask)
        if 'tasktype' in self.columnXref:
            self.doTaskTypeCell(self.currentTask, row.getCell(self.columnXref['tasktype']))
        if 'user' in self.columnXref:
            self.doUserCell(self.currentTask, row.getCell(self.columnXref['user']))

    def convertWorkbookToTemplate(self):
        print "%s\n" % self.templateName
        sheet = self.workbook.getSheetAt(0)
        # templateJson = {"id" : None, "type" : "xlrelease.Release", "title" : templateName, "phases" : [], "status" : "TEMPLATE"}
        template = Release()
        template.setTitle(self.templateName)
        template.setStatus(ReleaseStatus.TEMPLATE)
        template.setScheduledStartDate(Date())
        self.template = self.templateApi.createTemplate(template, self.targetFolderId)

        for row in sheet:
            print "Row %d\n" % row.getRowNum()
            if row.getRowNum() == 0:
                self.doHeaderRow(row)
            else:
                self.doDataRow(row)
