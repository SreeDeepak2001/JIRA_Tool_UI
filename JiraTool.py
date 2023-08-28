import datetime
import json
import sys
from configparser import ConfigParser
import requests
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import  QApplication, QPushButton, QMainWindow, QFileDialog, \
    QMessageBox, QTableWidgetItem, QTableWidget, QDialog, QHeaderView, QComboBox, QAction
from PyQt5.QtGui import QCursor, QIcon
from PyQt5 import uic
from jira import JIRA
from requests.auth import HTTPBasicAuth
from Dialog.configDialog import configDialog
from Dialog.aboutUsDialog import aboutUs
import webbrowser
import win32com.client as client
from Config.html_body import HTML_BODY
import pandas as pd
from Dialog.loadingScreen import Ui_Dialog

config = ConfigParser()
config.read('Config/config.ini')

api_token = config['JIRA_Config']['apitoken']
email = config['JIRA_Config']['email']
server = config['JIRA_Config']['server']



class Main(QMainWindow):

    updateprogress = pyqtSignal(int)
    progressfinished = pyqtSignal()
    searchstate = False

    def __init__(self):
        super(Main, self).__init__()
        self.load_ui()
        try:
            self.jira = JIRA(basic_auth=(email, api_token), server=server)
            self.config = ConfigParser()
        except Exception as e:
            self.warningBox(e)

    def load_ui(self):
        ids = self.getallprojects()
        uic.loadUi("Dialog/UI/form.ui",self)
        self.setWindowFlags(Qt.WindowCloseButtonHint | Qt.WindowMinimizeButtonHint)
        self.setWindowIcon(QIcon('Icon/icon.ico'))
        with open('Stylesheet/Combinear.qss') as f:
            qss = f.read()
        # self.setStyleSheet(qss)
        self.search = self.findChild(QPushButton, 'search')
        self.configuration = self.findChild(QPushButton, 'configuration')
        self.sendmail = self.findChild(QPushButton, 'sendmail')
        self.saveexcel = self.findChild(QPushButton, 'saveexcel')
        self.projectID = self.findChild(QComboBox, 'projectID')
        self.projectID.addItems(ids)
        self.about = self.findChild(QAction, "actionAbout_JIRA_Reminder_Tool")
        self.tablewidget = self.findChild(QTableWidget, "tableWidget")
        self.tablewidget.setColumnCount(6)
        header1 = QTableWidgetItem('Ticket Number')
        header2 = QTableWidgetItem('Priority')
        header3 = QTableWidgetItem('Assignee')
        header4 = QTableWidgetItem('Status')
        header5 = QTableWidgetItem('Last Comment Date')
        header0 = QTableWidgetItem('Select')
        self.tablewidget.setHorizontalHeaderItem(0, header0)
        self.tablewidget.setHorizontalHeaderItem(1, header1)
        self.tablewidget.setHorizontalHeaderItem(2, header2)
        self.tablewidget.setHorizontalHeaderItem(3, header3)
        self.tablewidget.setHorizontalHeaderItem(4, header4)
        self.tablewidget.setHorizontalHeaderItem(5, header5)
        self.header = self.tablewidget.horizontalHeader()
        self.header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        # header6 = QTableWidgetItem('Last Internal Comment')
        # header7 = QTableWidgetItem('Last External Comment')
        # self.tablewidget.setHorizontalHeaderItem(6, header6)
        # self.tablewidget.setHorizontalHeaderItem(7, header7)

        self.search.clicked.connect(lambda: self.searchclick2())
        self.configuration.clicked.connect(lambda: self.loadConfigDialog())
        self.sendmail.clicked.connect(lambda: self.sendMailclicked())
        self.saveexcel.clicked.connect(lambda: self.saveExcelClicked())



    def searchclick(self):
        self.searchstate = False
        projectID = self.projectID.currentText()
        if projectID:
            priority = self.getPriority()
            days_config = self.getdaysconfig()
            status = self.getStatus()
            if self.lastcommentExclude() is False:
                self.tablewidget.setColumnCount(8)
                header6 = QTableWidgetItem('Last Internal Comment')
                header7 = QTableWidgetItem('Last External Comment')
                self.tablewidget.setHorizontalHeaderItem(6, header6)
                self.tablewidget.setHorizontalHeaderItem(7, header7)
                self.header.setSectionResizeMode(6, QHeaderView.ResizeToContents)
                self.header.setSectionResizeMode(7, QHeaderView.ResizeToContents)
            else:
                self.tablewidget.setColumnCount(6)
            jquery = f"project = {projectID} {status}AND priority in ({priority}) ORDER BY priority"  # AND priority in (\"1 - Critical\",\"2 - Severe\")
            startat = 0
            self.tickets = []
            try:
                while True:
                    issues = self.jira.search_issues(jql_str=jquery, maxResults=1000, startAt=startat)
                    if len(issues) == 0:
                        self.warningBox('No Tickets found')
                        break
                    else:
                        self.addticketlist(issues, days_config)
                        startat = startat + 100
                        issues = self.jira.search_issues(jql_str=jquery, maxResults=1000, startAt=startat)
                        if len(issues) == 0:
                            break
                self.addtickettotable(self.tickets)
            except Exception as e:
                self.warningBox('Wrong Query')
                pass
        else:
            QApplication.restoreOverrideCursor()
            self.warningBox('No project Id found')

    def loadConfigDialog(self):
        try:
            self.loadconfig = configDialog()
            self.loadconfig.exec_()
        except Exception as e:
            self.warningBox(str(e))

    def warningBox(self,txt):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Warning)
        msg.setText(txt)
        msg.setWindowTitle("Warning")
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        msg.exec_()

    def getPriority(self):
        self.config.read('Config\config.ini')
        temp = []
        if self.config["Priority"]["critical"] == 'True':
            temp.append("1 - Critical")
        if self.config["Priority"]["severe"] == 'True':
            temp.append("2 - Severe")
        if self.config["Priority"]["moderate"] == 'True':
            temp.append("3 - Moderate")
        if self.config["Priority"]["minor"] == 'True':
            temp.append("4 - Minor")
        final = '\"1 - Critical\",\"2 - Severe\",\"3 - Moderate\",\"4 - Minor\"'
        if temp:
            while True:
                final = f"\"{temp[0]}\""
                if len(temp) > 1:
                    for t in temp[1:]:
                        final = final + ',' + f"\"{t}\""
                    break
                break
            return final

        else:
            self.warningBox("No priority selected")
            pass

    def getdaysconfig(self):
        self.config.read('config.ini')
        P1 = self.config['Days_Config']['critical']
        P2 = self.config['Days_Config']['severe']
        P3 = self.config["Days_Config"]['moderate']
        P4 = self.config["Days_Config"]["minor"]
        Days_Config = {
            "1 - Critical": P1,
            "2 - Severe": P2,
            "3 - Moderate": P3,
            "4 - Minor": P4
        }
        return Days_Config


    def checklastcomment(self, issue, commentConfig):
        """
            To get the last comments based on the comment Config
        :param issue: Issue ID (string)
        :param  commentConfig: Type of comment to check (string)
        :return: Date of the comment based on comment config or 0 (date object or 0)
        """
        global lastcommentdate
        comments = self.jira.comments(issue, expand="properties")
        if len(comments) != 0:
            comments.reverse()
            for comment in comments:
                if commentConfig == 'internal':
                    if not self.checkexternalcomment(issue, str(comment)):
                        lastcommentdate = comment.created
                        break
                    else:
                        lastcommentdate = 0
                elif commentConfig == 'external':
                    if self.checkexternalcomment(issue, str(comment)):
                        lastcommentdate = comment.created
                        break
                    else:
                        lastcommentdate = 0
                else:
                    lastcommentdate = comments[0].created
            if lastcommentdate is not None:
                return lastcommentdate
            else:
                return 0
        else:
            return 0

    @staticmethod
    def checkexternalcomment(issueKey, commentId):
        """
            gets the type of comment whether it's Internal or Public
        :param issueKey: Issue ID (string)
        :param commentId: Comment ID (string)
        :return: Type of comment ('True/False')
        """
        url = f"{server}/rest/api/3/issue/{issueKey}/comment/{commentId}"
        auth = HTTPBasicAuth(email, api_token)
        headers = {
            "Accept": "application/json",
        }
        response = requests.get(url, headers=headers, auth=auth)
        data = json.loads(response.text)
        return data['jsdPublic']

    def addticketlist(self, issues, CONFIG):
        """
            Add sorted tickets based on the CONFIG to a LIST
        :param issues: List of issues from JIRA API (list)
        :param CONFIG: CONFIG Dictionary for the project (dict)
        :return: List of Tickets. where, Ticket are the Dict of issue info formed based on the CONFIG given (list)
        """
        # logger.info('Segregating issues based on CONFIG')
        self.config.read('config.ini')
        if self.config["COMMENT_TYPE"]['internal'] == 'True':
            commenttype = 'internal'
        if self.config["COMMENT_TYPE"]['external'] == 'True':
            commenttype = 'external'
        if self.config["COMMENT_TYPE"]['all'] == 'True':
            commenttype = 'all'
        count = 1
        if issues:
            for issue in issues:
                print(self.searchstate)
                if self.searchstate:
                    break
                print(count)
                count = count+1
                lastCommentDate = self.checklastcomment(issue, commenttype)
                if lastCommentDate != 0:
                    UTCdate = self.getdate(lastCommentDate)
                    lastComment = self.formatdate(lastCommentDate)
                else:
                    checkDate = issue.fields.created
                    UTCdate = self.getdate(checkDate)
                    lastComment = "Not Commented Yet"

                if issue.fields.status != "More Information":
                    dateCheck = self.comparedate(UTCdate, issue.fields.priority, CONFIG)
                    priority_msg = str(issue.fields.priority)
                else:
                    if CONFIG.get("More Information") != "None":
                        dateCheck = self.comparedate(UTCdate, "More Information", CONFIG)
                        priority_msg = "More Information"
                    else:
                        dateCheck = 0
                if issue.fields.assignee is None:
                    assignee = "Unassigned"
                else:
                    assignee = issue.fields.assignee.displayName

                if self.lastcommentExclude() is False:
                    lastInternalComment = self.getinternalcomment(issue)
                    lastExternalComment = self.getexternalcomment(issue)
                    if dateCheck == 1:
                        ticket = {
                            "ticket_number": issue.key,
                            "priority": priority_msg,
                            "assignee": assignee,
                            "status": issue.fields.status,
                            "lastcomment": str(lastComment),
                            "lastInternalComment": str(lastInternalComment),
                            "lastExternalComment": str(lastExternalComment)
                        }
                        self.tickets.append(ticket)
                        self.updateprogress.emit(len(self.tickets))

                else:
                    if dateCheck == 1:
                        ticket = {
                            "ticket_number": issue.key,
                            "priority": priority_msg,
                            "assignee": assignee,
                            "status": issue.fields.status,
                            "lastcomment": str(lastComment)
                        }
                        self.tickets.append(ticket)
                        self.updateprogress.emit(len(self.tickets))


        else:
            self.warningBox('No Ticket found')

    @staticmethod
    def getdate(string):
        """
        converts STRING object to a DATE object in UTC time zone
        :param string: date (string)
        :return: Date object
        """
        year = int(string[0:4])
        month = int(string[5:7])
        day = int(string[8:10])
        hour = int(string[11:13])
        minute = int(string[14:16])
        seconds = int(string[17:19])
        TZD = string[23:28]
        date = datetime.datetime(year=year, month=month, day=day, hour=hour, minute=minute, second=seconds)
        if TZD == "-0800":
            final = date + datetime.timedelta(hours=8)
        else:
            final = date + datetime.timedelta(hours=7)
        return final

    def checklastcomment(self, issue, commentConfig):
        """
            To get the last comments based on the comment Config
        :param issue: Issue ID (string)
        :param  commentConfig: Type of comment to check (string)
        :return: Date of the comment based on comment config or 0 (date object or 0)
        """
        global lastcommentdate
        comments = self.jira.comments(issue, expand="properties")
        if len(comments) != 0:
            comments.reverse()
            for comment in comments:
                if commentConfig == 'internal':
                    if not self.checkexternalcomment(issue, str(comment)):
                        lastcommentdate = comment.created
                        break
                    else:
                        lastcommentdate = 0
                elif commentConfig == 'external':
                    if self.checkexternalcomment(issue, str(comment)):
                        lastcommentdate = comment.created
                        break
                    else:
                        lastcommentdate = 0
                else:
                    lastcommentdate = comments[0].created
            if lastcommentdate is not None:
                return lastcommentdate
            else:
                return 0
        else:
            return 0

    @staticmethod
    def comparedate(date, priority, CONFIG):
        """
        To Compare the last comment date with the current date
        :param date: date which needed to be compared (date object)
        :param priority: priority of the issue (Jira Object/String)
        :param CONFIG: CONFIG of the project (dict)
        :return: 1 or 0 based on the compared date results(0/1)
        """
        if CONFIG.get(str(priority)) != "None":
            days = int(CONFIG.get(str(priority)))
            checkDate = date + datetime.timedelta(days=days)
            today = datetime.datetime.utcnow().replace(microsecond=0)
            if checkDate.isoweekday() == 6:
                temp = (checkDate + datetime.timedelta(days=2)) - today
                if temp.days < 0:
                    return 1
                else:
                    return 0
            elif checkDate.isoweekday() == 7:
                temp = (checkDate + datetime.timedelta(days=1)) - today
                if temp.days < 0:
                    return 1
                else:
                    return 0
            else:
                temp = checkDate - today
                if temp.days < 0:
                    return 1
                else:
                    return 0
        else:
            return 0

    @staticmethod
    def formatdate(date):
        """
            To Change the Date Format for our convenience
            :param date: date(string)
        """
        year = int(date[0:4])
        month = int(date[5:7])
        day = int(date[8:10])
        formatedDate = f"{day}-{month}-{year}"
        return formatedDate

    def addtickettotable(self,tickets):
        issuecount = 0
        if tickets:
            for issue in tickets:
                self.tablewidget.setRowCount(issuecount + 1)
                selectItem = QTableWidgetItem()
                ticketitem = QTableWidgetItem(str(issue.get("ticket_number")))
                priorityitem = QTableWidgetItem(str(issue.get("priority")))
                assigneeitem = QTableWidgetItem(str(issue.get("assignee")))
                statusitem = QTableWidgetItem(str(issue.get('status')))
                commentitem = QTableWidgetItem(str(issue.get('lastcomment')))
                lastintenalcommentitem = QTableWidgetItem(str(issue.get('lastInternalComment')))
                lastextenalcommentitem = QTableWidgetItem(str(issue.get('lastExternalComment')))
                selectItem.setCheckState(Qt.Unchecked)
                self.tablewidget.setItem(issuecount, 0, selectItem)
                self.tablewidget.setItem(issuecount, 1, ticketitem)
                self.tablewidget.setItem(issuecount, 2, priorityitem)
                self.tablewidget.setItem(issuecount, 3, assigneeitem)
                self.tablewidget.setItem(issuecount, 4, statusitem)
                self.tablewidget.setItem(issuecount, 5, commentitem)
                self.tablewidget.setItem(issuecount, 6, lastintenalcommentitem)
                self.tablewidget.setItem(issuecount, 7, lastextenalcommentitem)
                issuecount = issuecount + 1
        self.progressfinished.emit()
        pass

    def reflink(self,index):
        if index.column() == 1:
            data = index.data()
            href = f"https://ServerName.atlassian.net/browse/{data}"
            webbrowser.open(href)
        pass

    def getcheckedtickets(self):
        columns = self.tablewidget.rowCount()
        selectedTickets = []
        final = []
        for index in range(columns):
            if self.tablewidget.item(index, 0).checkState() == Qt.Checked:
                selectedTickets.append(index)
        if selectedTickets:
            for issue in selectedTickets:
                ticket = {
                    "ticket_number": self.tablewidget.item(issue, 1).text(),
                    "priority": self.tablewidget.item(issue, 2).text(),
                    "assignee": self.tablewidget.item(issue, 3).text(),
                    "status": self.tablewidget.item(issue, 4).text(),
                    "lastcomment": self.tablewidget.item(issue, 5).text()
                }
                final.append(ticket)
        else:
            for i in range(columns):
                ticket = {
                    "ticket_number": self.tablewidget.item(i, 1).text(),
                    "priority": self.tablewidget.item(i, 2).text(),
                    "assignee": self.tablewidget.item(i, 3).text(),
                    "status": self.tablewidget.item(i, 4).text(),
                    "lastcomment": self.tablewidget.item(i, 5).text()
                }
                final.append(ticket)
        return final

    def sendoutlookmail(self, tableContent):
        """
        To Send Emails automatically via OutLook
        :param tableContent: List of Tickets need to be added to the MAIL(List)
        """
        try:
            outlook = client.Dispatch("Outlook.Application")
            message = outlook.CreateItem(0)
            attachment = message.Attachments.Add("D:\python examples\logo.png")
            attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "logo_img")
            message.Subject = "Jira Reminder Tool "
            message.Body = message.HTMLBody = HTML_BODY(tableContent)
            message.Display()
        except Exception as e:
            self.warningBox(str(e))
            pass

    def sendMailclicked(self):
        if self.tablewidget.rowCount() > 0:
            selected = self.getcheckedtickets()
            rows = self.gettable(selected)
            self.sendoutlookmail(rows)
        else:
            self.warningBox('No Tickets for the mail')

    @staticmethod
    def gettable(tickets):
        """
        takes the tickets and converts it to HTML Table format
        :param tickets: List of tickets (list)
        :return: HTML Table version of the list of tickets (list)
        :return: HTML Table version of the list of tickets (list)
        """
        count = 0
        rows = ""
        for ticket in tickets:
            ticket_number = ticket.get("ticket_number")
            priority = ticket.get("priority")
            assignee = ticket.get("assignee")
            status = ticket.get("status")
            lastcomment = ticket.get("lastcomment")
            href = f"https://ServerName.atlassian.net/browse/{ticket_number}"
            temp = f"<tr><td><a href = \"{href}\">{ticket_number}</a></td><td>{priority}</td><td>{assignee}</td>" \
                   f"<td>{status}</td><td>{lastcomment}</td></tr>"
            rows = rows + temp
            count = count + 1
        return rows

    def saveExcelClicked(self):
        if self.tablewidget.rowCount() > 0:
            tickets = self.getcheckedtickets()
            df = pd.DataFrame(data=tickets)
            filename = QFileDialog.getSaveFileName(caption='Save as', filter='.xlxs')
            save = filename[0]+filename[1]
            try:
                writer = pd.ExcelWriter(filename[0]+'.xlsx', engine='xlsxwriter')
                df.to_excel(writer, sheet_name='Sheet1')
                writer.save()
            except Exception as e:
                self.warningBox(str(e))
            pass
        else:
            self.warningBox('No ticket to convert to Excel')

    def getexternalcomment(self,issue):
        comments = self.jira.comments(issue, expand="properties")
        if len(comments) != 0:
            comments.reverse()
            externalcomment = 'No External Comments'
            for comment in comments:
                if self.checkexternalcomment(issue, str(comment)):
                    externalcommentdate = comment.created
                    externalcomment = self.formatdate(str(externalcommentdate))
                    break
            return externalcomment
        else:
            return 'Not Commented yet '
        pass

    def getinternalcomment(self,issue):
        comments = self.jira.comments(issue, expand="properties")
        if len(comments) != 0:
            comments.reverse()
            Internalcomment = 'No Internal Comments'
            for comment in comments:
                if not self.checkexternalcomment(issue, str(comment)):
                    Internalcommentdate = comment.created
                    Internalcomment = self.formatdate(str(Internalcommentdate))
                    break
            return Internalcomment
        else:
            return 'Not Commented yet '
        pass

    def getallprojects(self):
        url = f"{server}/rest/api/3/project"
        try:
            auth = HTTPBasicAuth(username=email, password=api_token)
            headers = {
                "Accept": "application/json",
            }
            response = requests.get(url, headers=headers, auth=auth)
            data = json.loads(response.text)
            ids = []
            for item in data:
                ids.append(item["key"])
            return ids
        except Exception as e:
            self.warningBox(str(e))


    def test(self):
        self.loadAboutUs = aboutUs()
        self.loadAboutUs.exec()

    def getStatus(self):
        config.read('Config/config.ini')
        if config["Status"]["status"] == "False":
            status = "AND status not in (Closed,Close,Complete,Completed) "
        else:
            status = ""
        return status

    @staticmethod
    def sendmail(userID, tableContent):
        try:
            outlook = client.Dispatch("Outlook.Application")
            message = outlook.CreateItem(0)
            message.To = userID
            attachment = message.Attachments.Add("Icon\logo.png")
            attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "logo_img")
            message.Subject = "Jira Reminder Tool "
            message.Body = message.HTMLBody = HTML_BODY(tableContent)
            message.Display()
        except Exception as e:
            print(e)
            pass

    def lastcommentExclude(self):
        config.read('Config/config.ini')
        if config["Days_Config"]["exclude"] == 'True':
            return True
        else:
            return False

    def searchclick2(self):
        projectID = self.projectID.currentText()
        if projectID:
            priority = self.getPriority()
            days_config = self.getdaysconfig()
            status = self.getStatus()
            if self.lastcommentExclude() is False:
                self.tablewidget.setColumnCount(8)
                header6 = QTableWidgetItem('Last Internal Comment')
                header7 = QTableWidgetItem('Last External Comment')
                self.tablewidget.setHorizontalHeaderItem(6, header6)
                self.tablewidget.setHorizontalHeaderItem(7, header7)
                self.header.setSectionResizeMode(6, QHeaderView.ResizeToContents)
                self.header.setSectionResizeMode(7, QHeaderView.ResizeToContents)
            else:
                self.tablewidget.setColumnCount(6)
            jquery = f"project = {projectID} {status}AND priority in ({priority}) ORDER BY priority"
            totalTicket = self.getTotalTickets(jquery)
            self.Dialog = QDialog()
            self.loading = Ui_Dialog()
            self.loading.setupUi(self.Dialog)
            self.loading.getMaximum(totalTicket)
            worker = WorkerThread(self)
            worker.start()
            self.updateprogress.connect(self.updateProgressBar)
            if self.Dialog.exec_() == 0 :
                worker.terminate()
                print("rejected")
                self.searchstate = True
            self.progressfinished.connect(lambda: self.closeProgressbar())

    def closeProgressbar(self):
        self.Dialog.close()


    def updateProgressBar(self, val):
        self.loading.getprogressval(val)

    def getTotalTickets(self,jquery):
        tots = 0
        startat = 0
        while True:
            issues = self.jira.search_issues(jql_str=jquery, maxResults=1000, startAt=startat)
            if len(issues) == 0:
                break
            else:
                startat = startat + 100
                tots = tots + len(issues)
                issues = self.jira.search_issues(jql_str=jquery, maxResults=1000, startAt=startat)
                if len(issues) == 0:
                    break
        return tots


class WorkerThread(QThread):

    def __init__(self, object):
        super(WorkerThread, self).__init__()
        self.obj = object

    def run(self):
        self.obj.searchclick()

        pass



if __name__ == "__main__":
    app = QApplication([])
    widget = Main()
    widget.getallprojects()
    widget.show()
    sys.exit(app.exec_())
