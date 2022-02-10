# # arcomp.py - Python program to compare the results of two autoruns executions to find changes in autoruns entries on Windows systems
# # Based on Autoruns for Widows by Mark Russinovich on Windows Sysinternals

# To Do
# - Better error handling
# Test syslog better

# Library imports
import configparser
from configparser import SafeConfigParser
import argparse
import os
import glob
import csv
import sqlite3
import json
import sys
from datetime import datetime
import subprocess
import smtplib
import email
import tempfile
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import time
import socket
import logging
from logging.handlers import SysLogHandler

# Global program info. Do Not Change.
version = ['1.0.0','Beta 1']
gitSourceUrl = 'https://github.com/HandyGuySoftware/arcomp'

def oops(msg):
    print(msg, file=sys.stderr)
    db.dbRollback()
    db.dbClose()
    exit(1)

class IniOptions:
    iniParser = None

    def __init__(self, iniPath):
        # First, see if the database is there. If not, need to create it
        self.openIniFile(iniPath);
        return None

    def openIniFile(self, iniFileSpec):
        try:
            self.iniParser = configparser.ConfigParser(interpolation=None)
            self.iniParser.read(iniFileSpec)
        except iniParser.ParsingError as err:
            return False
        return True

    def getIniOption(self, section, option):
        if self.iniParser.has_option(section, option):
            opt = self.iniParser.get(section, option)
            if opt != '':
                return opt
            else:
                return None
        else:
            return None

    def getIniSection(self, section):
        return dict(self.iniParser.items(section))

class Database:
    dbConn = None
    def __init__(self, dbPath):
        self.dbConn = sqlite3.connect(dbPath)   # Connect to database
        return None

    def dbClose(self):
        # Don't attempt to close a non-existant conmnection
        if self.dbConn:
            self.dbConn.close()
        self.dbConn = None
        return None

    # Clear database
    def dbSetup(self):

        # Get count of tables named 'history.' If the count is not 1, then the table doesn't exist, create it
        curs = self.dbConn.cursor()
        curs.execute("SELECT count(name) FROM sqlite_master WHERE type='table' AND name='history'")
        if curs.fetchone()[0]!=1:
            self.execSqlStmt('CREATE TABLE "history" ( `run_id` TEXT, `action` TEXT, `keyword` TEXT, `time` TEXT, `location` TEXT, `entry` TEXT, \
                `enabled` TEXT, `category` TEXT, `profile` TEXT, `description` TEXT, `signer` TEXT, `company` TEXT, `imagepath` TEXT, `version` TEXT, \
                `launchstring` TEXT, `vtdetection` TEXT, `vtpermalink` TEXT, `md5` TEXT, `sha1` TEXT, `pesha1` TEXT, `pesha256` TEXT, `sha256` TEXT, `imp` TEXT)')
        self.dbCommit()
        return None

    # Commit pending database transaction
    def dbCommit(self):
        if self.dbConn:     # Don't try to commit to a nonexistant connection
            self.dbConn.commit()
        return None

    def dbRollback(self):
        if self.dbConn:
            self.dbConn.rollback()
        return None

    # Execute a Sqlite command and manage exceptions
    # Return the cursor object to the command result
    def execSqlStmt(self, stmt, values = None):
        if not self.dbConn:
            return None

        # Set db cursor
        if values is None:
            curs = self.dbConn.cursor()
            curs.execute(stmt)
        else:
            self.dbConn.execute(stmt, values)
            curs = None
       
        return curs

    def getTableFieldNames(self, table):
        result = {}
        db.row_factory = sqlite3.Row
        curs = db.execSqlStmt("SELECT * FROM {}".format(table))
        rows = curs.fetchall()
        flds = [description[0] for description in curs.description]     # get field names

        return flds

def processCmdLineArgs():
    # Parse command line options with ArgParser library
    argParser = argparse.ArgumentParser(description='arcomp options.')

    argParser.add_argument("-f","--file", help="Specify the .csv file to load into system.", action="store")
    argParser.add_argument("-w","--write", help="Write output to a file. Format for argument is '-w <fname>,<type>'. Valid types are 'text', 'html', 'csv', and 'json'", action="append")
    argParser.add_argument("-e","--email", help="Send result to an email account. Make sure the [email] secition of te arcomp.ini file is filled in properly", action="store_true")
    argParser.add_argument("-s","--syslog", help="Send output to syslog server", action="store")
    argParser.add_argument("-x","--xfile", help="Hidden option: process an individual file manually", action="store")
    argParser.add_argument("-c", "--content", type=str, help="Specify sections to report ('a'dd, 'r'emove, or 's'ame)")
    argParser.add_argument("-r", "--runhistory", help="Print full execution history", action="store_true")
    argParser.add_argument("-R", "--runremove", help="Remove a specific <run_id> from the database.", action="store")
    try:
        cmdLineArgs = argParser.parse_args()
    except:
        oops("Command line parsing exception.")
    return cmdLineArgs

def loadAutoRunData(options):
    with open(options['file']) as f:
        reader = csv.reader(f)
        for row in reader:
            if row[0] == 'Time':  # Skip header row
                continue

            fldlist = ''
            vallist = ''
            for fld in options['dbfields']:
                fldlist += '{},'.format(fld)
                vallist += '?,'
            fldlist = fldlist[:-1]      # Remove trailing ','
            vallist = vallist[:-1]      # Remove trailing ','

            #sqlStmt = "INSERT INTO history ({}) VALUES (\'{}\', \'{}\', {})".format(fldlist, options['run_id'],row[1]+'-'+row[2], vallist)
            sqlStmt = "INSERT INTO history ({}) VALUES ({})".format(fldlist,vallist)

            rowTup = (options['run_id'],'', row[1]+'-'+row[2]) + tuple(row)
            if len(rowTup) < len(options['dbfields']): # need to pad fields
                for j in range(len(options['dbfields']) - len(rowTup)):
                    rowTup += ('',)
            db.execSqlStmt(sqlStmt, rowTup)
    return None

def getLastRunId():
    curs = db.execSqlStmt('SELECT DISTINCT run_id FROM history ORDER BY run_id DESC LIMIT 0,1')
    lastRunId = curs.fetchone()
    if lastRunId is None:       # Empty DB - no last run_id available
        return ''
    else:
        return lastRunId[0]

def compareAutoRunData(options):
    lastRunId = options['last_runid']

    # if last_runid == '', this is the first run. Everything gets added.
    if options['last_runid'] == '':
        curs = db.execSqlStmt("UPDATE history SET action='ADDED' WHERE run_id = '{}'".format(options['run_id']))
    else:
        curs = db.execSqlStmt("SELECT DISTINCT keyword FROM history WHERE run_id == '{}' and keyword NOT IN (SELECT DISTINCT keyword FROM history WHERE run_id == '{}')".format(options['run_id'], options['last_runid']))
        distinctRows = curs.fetchall()
        if len(distinctRows) != 0:
            for keyword in distinctRows:
                curs = db.execSqlStmt("UPDATE history SET action='ADDED' where run_id='{}' and keyword='{}'".format(options['run_id'],keyword[0]))

    # See what was deleted since the last run
    # HOWEVER, if something was deleted in the last run, a 'REMOVED' record was added to that run. 
    #   The record still won't show up in this run and will again have another 'REMOVED' record added, unless we strp in to stop this tragedy.
    if options['last_runid'] != '':             # if last_runid == '', this is the first run. There's nothing that can be deleted.
        curs = db.execSqlStmt("CREATE TEMPORARY TABLE tmphistory AS SELECT * FROM history WHERE run_id == '{}' AND action != 'REMOVED' \
            AND keyword NOT IN \
                (SELECT DISTINCT keyword FROM history WHERE run_id == '{}')".format(options['last_runid'], options['run_id']))
        curs = db.execSqlStmt("UPDATE tmphistory SET run_id='{}', action='REMOVED'".format(options['run_id']))
        curs = db.execSqlStmt("INSERT INTO history SELECT * FROM tmphistory")
        curs = db.execSqlStmt("DROP TABLE IF EXISTS tmphistory")

    # See what's the same since the last run
    if options['last_runid'] != '':             # if last_runid == '', this is the first run. There's nothing that's the same.
        curs = db.execSqlStmt("UPDATE history SET action='SAME' WHERE run_id='{}' and action IS ''".format(options['run_id']))
    return

def generateDictFromSql(sql):
    result = {}
    db.row_factory = sqlite3.Row
    curs = db.execSqlStmt(sql)
    flds = [description[0] for description in curs.description]     # get field names

    for r in curs:
       result[r[2]] = {}
       for i in range(len(flds)):
            result[r[2]][flds[i]] = r[i]

    return flds, result

def generateReport(options):
    rptOutput = {
        'added': {
            'name': 'added',
            'title': 'Entries Added Since Last Run',
            }, 
        'removed': {
            'name': 'removed',
            'title': 'Entries Removed Since Last Run',
            }, 
        'same': {
            'name': 'same',
            'title': 'Entries Unchanged Since Last Run',
            }
        }
    
    rptOutput['added']['fieldnames'], rptOutput['added']['result'] = generateDictFromSql("SELECT * FROM history WHERE run_id='{}' and action='ADDED'".format(options['run_id']))
    rptOutput['removed']['fieldnames'], rptOutput['removed']['result'] = generateDictFromSql("SELECT * FROM history WHERE run_id='{}' and action='REMOVED'".format(options['run_id']))
    rptOutput['same']['fieldnames'], rptOutput['same']['result'] = generateDictFromSql("SELECT * FROM history WHERE run_id='{}' and action='SAME'".format(options['run_id']))

    return rptOutput

def buildHTML(data, options):

    html = "<table border=1>"
    if 'a' in options['content']:
        html += "<tr><td colspan = {} align=center> <b>Entries Added</b></td></tr>\n".format(len(options['reportfields']))
        if len(data['added']['result']) == 0:
            html += "<tr><td colspan = {} align=center>(None)</td></tr>".format(len(options['reportfields']))
        else:
            html+= "<tr>"
            for i in range(len(data['added']['fieldnames'])):
                if data['added']['fieldnames'][i] in options['reportfields']:
                    html += "<th>{}</th>".format(data['added']['fieldnames'][i])
            html += "</tr>\n"
            for key, values in data['added']['result'].items():
                html += "<tr>"
                for i in range(len(data['added']['fieldnames'])):
                    if data['added']['fieldnames'][i] in options['reportfields']:
                        html += '<td>{}</td>'.format(values[data['added']['fieldnames'][i]])
                html += '</tr>\n'
   
    if 'r' in options['content']:
        html += "<tr><td colspan = {} align=center> <b>Entries Removed</b></td></tr>\n".format(len(options['reportfields']))
        if len(data['removed']['result']) == 0:
            html += "<tr><td colspan = {} align=center>(None)</td></tr>".format(len(options['reportfields']))
        else:
            html+= "<tr>"
            for i in range(len(data['removed']['fieldnames'])):
                if data['removed']['fieldnames'][i] in options['reportfields']:
                    html += "<th>{}</th>".format(data['removed']['fieldnames'][i])
            html += '</tr>\n'
            for key, values in data['removed']['result'].items():
                html += "<tr>"
                for i in range(len(data['removed']['fieldnames'])):
                    if data['removed']['fieldnames'][i] in options['reportfields']:
                        html += '<td>{}</td>'.format(values[data['removed']['fieldnames'][i]])
                html += '</tr>\n'
   
    if 's' in options['content']:
        html += "<tr><td colspan = {} align=center> <b>Entries Unchanged</b></td></tr>\n".format(len(options['reportfields']))
        if len(data['same']['result']) == 0:
            html += "<tr><td colspan = {} align=center>(None)</td></tr>".format(len(options['reportfields']))
        else:
            html+= "<tr>"
            for i in range(len(data['same']['fieldnames'])):
                if data['same']['fieldnames'][i] in options['reportfields']:
                    html += "<th>{}</th>".format(data['same']['fieldnames'][i])
            html += '</tr>\n'
            for key, values in data['same']['result'].items():
                html += "<tr>"
                for i in range(len(data['same']['fieldnames'])):
                    if data['same']['fieldnames'][i] in options['reportfields']:
                        html += '<td>{}</td>'.format(values[data['same']['fieldnames'][i]])
                html += '</tr>\n'
    html += '</table>\n'
    html+= '<br><br>Report generated by <a href="{}">arcomp</a> version {} ({})<br>'.format(gitSourceUrl, version[0], version[1])
    return html

def buildText(data, options):

    if 'a' in options['content']:
        text = "Entries Added\n"
        for i in range(len(data['added']['fieldnames'])):
            if data['added']['fieldnames'][i] in options['reportfields']:
                text += data['added']['fieldnames'][i] + ' | '
        text += '\n'
        for key, values in data['added']['result'].items():
            for i in range(len(data['added']['fieldnames'])):
                if data['added']['fieldnames'][i] in options['reportfields']:
                    text += '{} |'.format(values[data['added']['fieldnames'][i]])
            text += '\n'

    if 'r' in options['content']:
        text += "Entries Removed\n"
        for i in range(len(data['removed']['fieldnames'])):
            if data['removed']['fieldnames'][i] in options['reportfields']:
                text += data['removed']['fieldnames'][i] + ' | '
        text += '\n'
        for key, values in data['removed']['result'].items():
            for i in range(len(data['removed']['fieldnames'])):
                if data['removed']['fieldnames'][i] in options['reportfields']:
                    text += '{} |'.format(values[data['removed']['fieldnames'][i]])
            text += '\n'

    if 's' in options['content']:
        text += "Entries Unchanged\n"
        for i in range(len(data['same']['fieldnames'])):
            if data['same']['fieldnames'][i] in options['reportfields']:
                text += data['same']['fieldnames'][i] + ' | '
        text += '\n'
        for key, values in data['same']['result'].items():
            for i in range(len(data['same']['fieldnames'])):
                if data['same']['fieldnames'][i] in options['reportfields']:
                    text += '{} |'.format(values[data['same']['fieldnames'][i]])
            text += '\n'

    text += '\nReport generated by arcomp ({}) Version {} ({})\n'.format(gitSourceUrl, version[0], version[1])
    return text

def buildCSV(data, options):
    text=''
    
    if 'a' in options['content']:
        text = "Entries Added,\n"
        for i in range(len(data['added']['fieldnames'])):
            if data['added']['fieldnames'][i] in options['reportfields']:
                text += '"{}",'.format(data['added']['fieldnames'][i])
        text += '\n'
        for key, values in data['added']['result'].items():
            for i in range(len(data['added']['fieldnames'])):
                if data['added']['fieldnames'][i] in options['reportfields']:
                    text += '"{}",'.format(values[data['added']['fieldnames'][i]])
            text += '\n'

    if 'r' in options['content']:
        text += "Entries Removed,\n"
        for i in range(len(data['removed']['fieldnames'])):
            if data['removed']['fieldnames'][i] in options['reportfields']:
                text += '"{}",'.format(data['removed']['fieldnames'][i])
        text += '\n'
        for key, values in data['removed']['result'].items():
            for i in range(len(data['removed']['fieldnames'])):
                if data['removed']['fieldnames'][i] in options['reportfields']:
                    text += '"{}",'.format(values[data['removed']['fieldnames'][i]])
            text += '\n'

    if 's' in options['content']:
        text += "Entries Unchanged,\n"
        for i in range(len(data['same']['fieldnames'])):
            if data['same']['fieldnames'][i] in options['reportfields']:
                text += '"{}",'.format(data['same']['fieldnames'][i])
        text += '\n'
        for key, values in data['same']['result'].items():
            for i in range(len(data['same']['fieldnames'])):
              if data['same']['fieldnames'][i] in options['reportfields']:
                    text += '"{}",'.format(values[data['same']['fieldnames'][i]])
            text += '\n'

    text += '\nReport generated by arcomp ({}) Version {} ({})\n'.format(gitSourceUrl, version[0], version[1])
    return text

def writeFiles(data, options):
    for item in options['write'].items():
        if item[1] == 'text':                   # Convert to text
            output = buildText(data, options)
        elif item[1] == 'html':                 # Convert to HTML
            output = buildHTML(data, options)
        elif item[1] == 'csv':                  # Convert to CSV
            output = buildCSV(data, options)
        else:                                   # Data is already in JSON format internally
            output = data

        outfile = open(item[0],'w')
        if item[1].lower() in ['text','html','csv']:   # Write data to file    
            outfile.write(output)
        else:                                   # Json data requires special call
            json.dump(output, outfile)
        outfile.close()

def sendEmail(data, options, inifile):
    try:
        serverconnect = smtplib.SMTP(options['email']['server'],options['email']['port'])
        if options['email']['encryption'] != None:   # Do we need to use SSL/TLS?
            try:
                tlsContext = ssl.create_default_context()
                serverconnect.starttls(context=tlsContext)
            except Exception as e:
                exit(1)
        try:
            pw = iniFile.getIniOption('email','password')
            retVal, retMsg = serverconnect.login(options['email']['account'], pw)  # Get password live so it's not stored in memory long-term
        except:
            e = sys.exc_info()[0]
    except (smtplib.SMTPAuthenticationError, smtplib.SMTPConnectError, smtplib.SMTPSenderRefused):
        e = sys.exc_info()[0]

    # Build email message
    msg = MIMEMultipart('alternative')
    msg['Subject'] = options['email']['subject']
    msg['From'] = options['email']['sender']
    msg['To'] = options['email']['receiver']
    msg['Date'] = email.utils.formatdate(time.time(), localtime=True)

    body= buildHTML(data, options)
    msgPart = MIMEText(body, 'html')
    msg.attach(msgPart)

    afilename = options['datapath'] + '\\arcompattachment.html'
    attachfile = open(afilename, "w")
    attachfile.write(body)
    attachfile.close()

    attachfile = open(afilename, "rb")
    part = MIMEBase('application','octet-stream')
    part.set_payload(attachfile.read())
    part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(afilename))
    encoders.encode_base64(part)
    msg.attach(part)
    attachfile.close()

    # Send the email
    serverconnect.send_message(msg, options['email']['sender'], options['email']['receiver'])
    return None

def sendSyslog(data, options):
    try:
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        handler = SysLogHandler(address=(options['syslog']['server'], options['syslog']['port']), facility = 16)
        logger.addHandler(handler)
    except :
        e = sys.exc_info()[0]

    now = datetime.now().isoformat()
    for key, values in data['added']['result'].items():
        logmsg = '[{}][{}][INFO][{}][ADDED]{}|{}|{}|{}|{}|{}|{}'.format(now, options['hostname'], options['run_id'],
            values['location'],values['entry'],values['description'],values['signer'],values['company'],values['imagepath'],values['launchstring'])
        logger.info(logmsg)
    # Send logmsg to syslog

    for key, values in data['removed']['result'].items():
        logmsg = '[{}][{}][INFO][{}][REMOVED]{}|{}|{}|{}|{}|{}|{}'.format(now, options['hostname'], options['run_id'],
            values['location'],values['entry'],values['description'],values['signer'],values['company'],values['imagepath'],values['launchstring'])
        # Send logmsg to syslog
        logger.info(logmsg)

    for key, values in data['same']['result'].items():
        logmsg = '[{}][{}][INFO][{}][SAME]{}|{}|{}|{}|{}|{}|{}'.format(now, options['hostname'], options['run_id'],
            values['location'],values['entry'],values['description'],values['signer'],values['company'],values['imagepath'],values['launchstring'])
        # Send logmsg to syslog
        logger.info(logmsg)

    handler.close
    return None

def printHistory():
    curs = db.execSqlStmt("SELECT DISTINCT run_id FROM history ORDER BY run_id ASC")
    runids = curs.fetchall()
    for i in range(len(runids)):
        id = runids[i][0]
        print("{}   ({}-{}-{}  {}:{}:{}.{})".format(id, id[0:4], id[4:6], id[6:8], id[9:11], id[11:13], id[13:15], id[16:]))
    return

def deleteRunID(runid):
    curs = db.execSqlStmt("SELECT run_id FROM history WHERE run_id = '{}'".format(runid))
    result = curs.fetchall()
    if len(result) == 0:
        print("No such run_id: {}".format(runid))
        return None

    curs = db.execSqlStmt("DELETE FROM history WHERE run_id = '{}'".format(runid))
    result = curs.fetchall()

    return None

##### Let's Go! #####
if __name__ == "__main__":

    # Global options dictionary holds all the program options, either default, from the .ini file, or from the commmand line
    options = {}
    options['progpath'] = os.path.dirname(os.path.realpath(sys.argv[0]))  # Get program home directory
    options['hostname'] = socket.gethostname()
    options['run_id'] = datetime.now().strftime("%Y%m%d-%H%M%S-%f") # Each run gets its own run identifier. All entries from the same run have the same run_id.
    options['version'] = version
    options['gitSourceUrl'] = gitSourceUrl

    # Open and read the .ini file
    iniFile = IniOptions(options['progpath'] + '\\arcomp.ini')              # Class to handle .ini file operations
    options['autorunspath'] = iniFile.getIniOption('main','autorunspath')    # Path to autorunsc.exe. If Null, assume it's in thre Windows %PATH%
    options['datapath'] = iniFile.getIniOption('main','datapath')            # Path to data files. If Null, assume it's in the same directory as this program
    if options['datapath'] is None:                                         # Use default data path
        options['datapath'] = options['progpath']
    options['reportfields'] = list({key: value for key, value in iniFile.getIniSection('fields').items() if value == 'True'})     # List of fields to use in report output. From .ini file [report] section

    # Get and process command line arguments
    progArgs = processCmdLineArgs()
    options['file'] = progArgs.file

    if progArgs.write is not None:      # output files specified.
        options['write'] = {}
        for i in range(len(progArgs.write)):
            fname,type = progArgs.write[i].split(",")                   # Exception check Here
            if type.lower() not in ['text','html','csv','json']:
                print("Command line error, -w option. Invalid type: {}. Filetype must be 'text', 'html', 'csv', or 'json'".format(type))
                exit(1)
            options['write'][fname] = type.lower()
    
    if progArgs.syslog is not None:     # Output to syslog specified
        options['syslog'] = {}
        syslogspec = progArgs.syslog.split(':')
        options['syslog']['server'] = syslogspec[0]
        if len(syslogspec) == 1:                # No port specified. Use default of 514
            options['syslog']['port'] = 514
        else:
            options['syslog']['port'] = int(syslogspec[1])

    if progArgs.content is None:
        options['content'] = 'ars'
    else:
        for i in range(len(progArgs.content)):
            if progArgs.content[i] not in ['a','r','s']:
                print("--content option: invalid option: '{}'. Must be a combination of 'a', 'r', and/or 's'".format(progArgs.content[i]))
                exit(1)
        options['content'] = progArgs.content

    # Check for email options
    options['email'] = {}
    options['email']['send'] = progArgs.email
    if options['email']['send'] is True:
        options['email']['server'] = iniFile.getIniOption('email','server')  
        options['email']['port'] = iniFile.getIniOption('email','port')  
        options['email']['encryption'] = iniFile.getIniOption('email','encryption')  
        options['email']['account'] = iniFile.getIniOption('email','account')  
        #options['email']['password'] = iniFile.getIniOption('email','password')  
        options['email']['sender'] = iniFile.getIniOption('email','sender')  
        options['email']['sendername'] = iniFile.getIniOption('email','sendername')  
        options['email']['receiver'] = iniFile.getIniOption('email','receiver')  
        options['email']['authentication'] = iniFile.getIniOption('email','authentication') 
        options['email']['subject'] = iniFile.getIniOption('email','subject') 

    # Open and prep database
    db = Database(options['datapath'] + '\\arcompdata.db')
    db.dbSetup()
    options['dbfields'] = db.getTableFieldNames('history')
    
    # Need to just print history?
    if progArgs.runhistory is True:
        printHistory()
        exit(0)

    # Need to delete a run_id?
    if progArgs.runremove is not None:
        deleteRunID(progArgs.runremove)
        db.dbCommit()
        db.dbClose()
        exit(0)

    # Get last run_id
    options['last_runid'] = getLastRunId()

    # Are we processing a command-line file or letting autorunsc.exe do its thing?
    if options['file'] is None:                 # There's no specific file to process. Execute autorunsc.exe and collect output file
        cmdline = '\"\"{}\" -a * -c -h -s -u -v -vt -o \"\"{}\\aroutput.csv\"'.format(options['autorunspath'],options['datapath'])  
        result = os.system(cmdline)
        options['file'] = '{}\\aroutput.csv'.format(options['datapath'])

    # Load data from file
    loadAutoRunData(options)

    # Compare current run to last run and add results to database
    compareAutoRunData(options)

    # Generate report based on database results
    reportData = generateReport(options)

    # Do we need to send output to files?
    if 'write' in options:
        writeFiles(reportData, options)

    # Do we need to send email?
    if options['email']['send'] is True:
        sendEmail(reportData, options,iniFile)

    # Do we need to send to syslog?
    if progArgs.syslog is not None:
        sendSyslog(reportData, options)

    # Close database and exit``
    db.dbCommit()
    db.dbClose()

    exit(0)
