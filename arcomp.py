######
#
# Program name: arcomp.py
# Purpose:      Compare the results of two Autoruns executions to find changes in Autoruns entries on Windows systems
#               Based on Autoruns for Widows by Mark Russinovich on Windows Sysinternals
# Author:       Stephen Fried for Handy Guy Software
# 
#####

# Import system modules
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
version = ['1.0.1','Beta 2']
gitSourceUrl = 'https://github.com/HandyGuySoftware/arcomp'
copyright = 'Copyright (c) 2022 Stephen Fried for Handy Guy Software. Released under MIT license. See LICENSE file for details'

# Global error and exit handler
def oops(msg):
    sys.stderr.write(msg)
    if 'db' in vars():
        db.dbRollback()
        db.dbClose()
    if 'progLog' in vars():
        progLog.logClose()
    exit(1)

class Logger:
    logfile = None

    def __init__(self, fname):
        try:
            self.logfile = open(fname,'a')
            self.logWrite("Arcomp logfile - open")
        except (OSError, IOError):
            e = sys.exc_info()[0]
            sys.stderr.write('Error opening log file {}: {}\n'.format(fname, e))
            oops("Log file open error")
        return None

    def logWrite(self, msg):
        if self.logfile is not None:
            self.logfile.write('[{}][run_id:{}][DEBUG] {}\n'.format(datetime.now().isoformat(), options['run_id'], msg))
            self.logfile.flush()
        return None

    def logClose(self):
        if self.logfile is not None:
            self.logfile.close()
        self.logfile = None
        return None

# Class to manage .ini file handling
class IniOptions:
    iniParser = None

    def __init__(self, iniPath):
        # First, see if the .ini file is there. If not, need to create it
        self.openIniFile(iniPath);
        return None

    def openIniFile(self, iniFileSpec):
        try:
            self.iniParser = configparser.ConfigParser(interpolation=None, allow_no_value=True)
            self.iniParser.optionxform = str         # Set to read values in .ini file as case-sensitive
            self.iniParser.read(iniFileSpec)
        except self.iniParser.ParsingError as err:
            return False
        
        return True
    
    # Retrieve an individial value from '[secton] option='
    def getIniOption(self, section, option):
        if self.iniParser.has_option(section, option):
            opt = self.iniParser.get(section, option)
            if opt != '':
                return opt
            else:
                return None
        else:
            return None

    # Return an entire [section] from the .ini file as a dictionary
    def getIniSection(self, section):
        return dict(self.iniParser.items(section))

    def hasSection(self, section):
        if self.iniParser.has_section(section):
            return True
        else:
            return False

# SQLite database management class
class Database:
    dbConn = None   # atabase connection
    def __init__(self, dbPath):
        self.dbConn = sqlite3.connect(dbPath)   # Connect to database
        return None

    def dbClose(self):
        # Don't attempt to close a non-existant conmnection
        if self.dbConn:
            self.dbConn.close()
        self.dbConn = None
        return None

    # Initialize database, if needed
    def dbSetup(self):
        # Get count of tables named 'history.' If the count is not 1, then the table doesn't exist, so create it
        curs = self.dbConn.cursor()
        curs.execute("SELECT count(name) FROM sqlite_master WHERE type='table' AND name='history'")
        
        # (Re)build the history table. The fields here need to be named specifically, 
        #   as there is no way to automatically extract field names from a table that does not yet exist (see comment in self.getTableFieldNames()
        if curs.fetchone()[0] !=1:
            self.execSqlStmt('CREATE TABLE "history" ( `run_id` TEXT, `action` TEXT, `keyword` TEXT, `time` TEXT, `location` TEXT, `entry` TEXT, \
                `enabled` TEXT, `category` TEXT, `profile` TEXT, `description` TEXT, `signer` TEXT, `company` TEXT, `imagepath` TEXT, `version` TEXT, \
                `launchstring` TEXT, `vtdetection` TEXT, `vtpermalink` TEXT, `md5` TEXT, `sha1` TEXT, `pesha1` TEXT, `pesha256` TEXT, `sha256` TEXT, `imp` TEXT)')
        self.dbCommit()
        return None

    # Commit pending database transactions
    def dbCommit(self):
        if self.dbConn:     # Don't try to commit to a nonexistant connection
            self.dbConn.commit()
        return None

    # Rollback database transactions. 
    # Used in case there is an error in processing, so we can leave the database the way we found it.
    def dbRollback(self):
        if self.dbConn:
            self.dbConn.rollback()
        return None

    # Execute a SQLite command and manage exceptions
    # Return the cursor object to the command result
    # stmt = SQL statement to execute
    # values = tuple to use if the stmt is in the form 'UPDATE table (flds,...) VALUES (?, ?, ?....)
    def execSqlStmt(self, stmt, values = None):
        if not self.dbConn:     # Don't execute against a non-existant db connection
            return None
        # Set db cursor
        if values is None:                  # Somple SQL statement
            curs = self.dbConn.cursor()
            curs.execute(stmt)
        else:                               # Values-based update
            self.dbConn.execute(stmt, values)
            curs = None
        return curs

    # Retrieve the field names from a specific table
    # This is used so that the code does not have to be manually updated in the event the field configuration changes
    # Except that the fields DO need to be manually updated in self.dbSetup(), as you can't extract fields from a table that doesn't exist.
    def getTableFieldNames(self, table):
        result = {}
        db.row_factory = sqlite3.Row
        curs = db.execSqlStmt("SELECT * FROM {}".format(table))
        rows = curs.fetchall()
        flds = [description[0] for description in curs.description]     # get field names

        return flds

# Process command line arguments
def processCmdLineArgs():
    # Parse command line options with ArgParser library
    argParser = argparse.ArgumentParser(description='arcomp options.')

    argParser.add_argument("-c","--content", type=str, help="Specify sections to include in the report ('a'dd, 'r'emove, or 's'ame)")
    argParser.add_argument("-e","--email", help="Send report to an email account. Make sure the [email] section of the arcomp.ini file is filled in properly.", action="store_true")
    argParser.add_argument("-f","--file", help="Specify a .csv file to load into system. Must be created using 'autorunsc.exe -a * -c -h -s -u -v -vt -o <filename>'", action="store")
    argParser.add_argument("-r", "--runhistory", help="Print full history of autorunsc results.", action="store_true")
    argParser.add_argument("-R", "--runremove", help="Remove a specific <run_id> from the database.", action="store")
    argParser.add_argument("-s","--syslog", help="Send output to syslog server. Format is '-s <IP address or DNS name>[:port]'. Default port is 514", action="store")
    argParser.add_argument("-w","--write", help="Write report output to a file. Format for argument is '-w <fname>,<type>'. Valid types are 'text', 'html', 'csv', and 'json'", action="append")
    try:
        cmdLineArgs = argParser.parse_args()
    except:
        oops("Command line parsing exception.")
    return cmdLineArgs

# Load data from AutoRuns execution and add it to the database
def loadAutoRunData(options):
    progLog.logWrite('Loading Autoruns data from file {}'.format(options['file']))
                     
    # Load data lines from file, in .csv format
    with open(options['file']) as f:
        reader = csv.reader(f)
        for row in reader:
            if row[0] == 'Time':  # Skip header row
                continue

            # Create the field and values list for the impending SQLite INSERT call.
            fldlist = ''
            vallist = ''
            for fld in options['dbfields']:     # Run through each field in the history table
                fldlist += '{},'.format(fld)
                vallist += '?,'
            fldlist = fldlist[:-1]      # Remove trailing ','
            vallist = vallist[:-1]      # Remove trailing ','

            sqlStmt = "INSERT INTO history ({}) VALUES ({})".format(fldlist,vallist)

            # Create the tuple needed for the 'VALUES' section of the SQL statement
            # The 'keyword' field in the table is a unique key, a concatenation of the 'location' and 'entry' fields
            rowTup = (options['run_id'],'', row[1]+'-'+row[2]) + tuple(row)
            if len(rowTup) < len(options['dbfields']): # need to pad fields
                for j in range(len(options['dbfields']) - len(rowTup)):
                    rowTup += ('',)
            progLog.logWrite("Inserting new record: [{}]".format(row[1]+'-'+row[2]))
            db.execSqlStmt(sqlStmt, rowTup)
    return None

# Get the last run_id stored in the system. This is used to extract data from the last run to compare against the current run
def getLastRunId():
    curs = db.execSqlStmt('SELECT DISTINCT run_id FROM history ORDER BY run_id DESC LIMIT 0,1')
    lastRunId = curs.fetchone()
    progLog.logWrite("Retrieved last run_id: [{}]".format(lastRunId))
    if lastRunId is None:       # Empty DB - no last run_id available
        return ''
    else:
        return lastRunId[0]

# This is where the sausage is made. Run comparisions between the current run and last run data, looking for what's been added, removed, and left the same
def compareAutoRunData(options):
    progLog.logWrite("Comparing Autorun data: [{}] vs [{}]".format(options['run_id'], options['last_runid']))
    lastRunId = options['last_runid']

    # if last_runid == '', this is the first run. Everything gets added.
    if options['last_runid'] == '':
        progLog.logWrite("No last_runid. First time run. Everything gets added.")
        curs = db.execSqlStmt("UPDATE history SET action='ADDED' WHERE run_id = '{}'".format(options['run_id']))
    else:
        progLog.logWrite("Noting ADDED entries.")
        # Get rows where an entry is in the current run but not in the last run
        curs = db.execSqlStmt("SELECT DISTINCT keyword FROM history WHERE run_id == '{}' and keyword NOT IN (SELECT DISTINCT keyword FROM history WHERE run_id == '{}')".format(options['run_id'], options['last_runid']))
        distinctRows = curs.fetchall()
        if len(distinctRows) != 0:
            for keyword in distinctRows:
                # Set the action for that row to 'ADDED'
                curs = db.execSqlStmt("UPDATE history SET action='ADDED' where run_id='{}' and keyword='{}'".format(options['run_id'],keyword[0]))

    # See what was deleted since the last run (an entry is in the last run but not in the current run)
    # HOWEVER, if something was detected as deleted in the last run, a 'REMOVED' record was added to that run, creating a phantom record for an item that really wasn't found during the run.
    #   That new REMOVED record will show up here in last_run, but not in the current_run. Normally this would generate another REMOVED record, unless we step in to stop the madness.
    if options['last_runid'] != '':             # if last_runid == '', this is the first run. There's nothing that can be deleted.
        progLog.logWrite("Noting REMOVED entries.")
        # Create a temporary table to hold the changed data, then copy that table back into the history table.
        curs = db.execSqlStmt("CREATE TEMPORARY TABLE tmphistory AS SELECT * FROM history WHERE run_id == '{}' AND action != 'REMOVED' \
            AND keyword NOT IN \
                (SELECT DISTINCT keyword FROM history WHERE run_id == '{}')".format(options['last_runid'], options['run_id']))
        curs = db.execSqlStmt("UPDATE tmphistory SET run_id='{}', action='REMOVED'".format(options['run_id']))
        curs = db.execSqlStmt("INSERT INTO history SELECT * FROM tmphistory")
        curs = db.execSqlStmt("DROP TABLE IF EXISTS tmphistory")

    # See what's the same since the last run. Basically, whatever is not tagged as 'ADDED' or 'REMOVED' is tagged as 'SAME'.
    if options['last_runid'] != '':             # if last_runid == '', this is the first run. There's nothing that's the same.
        progLog.logWrite("Noting SAME entries.")
        curs = db.execSqlStmt("UPDATE history SET action='SAME' WHERE run_id='{}' and action IS ''".format(options['run_id']))
    return

# Generate a dictionary from a list of fields returned form an SQL query
def generateDictFromSql(sql):
    progLog.logWrite("generateDictFromSql(\'{}\')".format(sql))

    finalResult = {}        # Final values from SQL query
    
    db.row_factory = sqlite3.Row
    curs = db.execSqlStmt(sql)
    dbFlds = [description[0] for description in curs.description]     # get field names
 
    # Get index of 'signer' & 'company' fields
    signerIndex = dbFlds.index('signer')
    companyIndex = dbFlds.index('company')

    # Get the field number of the 'key' field
    keyFieldIndex = dbFlds.index('keyword')

    # Loop through db results
    for resultRow in curs:
        # If signer is on the ignore_signer list or company is on the ignore_company list, skip this row
        if resultRow[signerIndex] in options['ignore_signer'] or resultRow[companyIndex] in options['ignore_company']:
            progLog.logWrite("Skipping row. Key:[{}] signer:[{}]  company:[{}]".format(resultRow[keyFieldIndex], resultRow[signerIndex], resultRow[companyIndex]))
            continue

        finalResult[resultRow[keyFieldIndex]] = {}      
        for i in range(len(dbFlds)):
            finalResult[resultRow[keyFieldIndex]][dbFlds[i]] = resultRow[i]

    return dbFlds, finalResult

# Generate dictionaries for the three types of results ('ADDED', 'REMOVED', and 'SAME')
def generateReport(options):
    itemCount = getRunIdCount(options['run_id'])
    
    progLog.logWrite("Generating reports.")
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
    
    # For each type of result, select from the history table with the current run_id and action='<whatever>'
    rptOutput['added']['fieldnames'], rptOutput['added']['result'] = generateDictFromSql("SELECT * FROM history WHERE run_id='{}' and action='ADDED'".format(options['run_id']))
    rptOutput['removed']['fieldnames'], rptOutput['removed']['result'] = generateDictFromSql("SELECT * FROM history WHERE run_id='{}' and action='REMOVED'".format(options['run_id']))
    rptOutput['same']['fieldnames'], rptOutput['same']['result'] = generateDictFromSql("SELECT * FROM history WHERE run_id='{}' and action='SAME'".format(options['run_id']))
    return rptOutput

# Build an HTML-style output
# The resulting report output is modified based on the select of desired fields in the [fields] section of the .ini file
# NOTE: comments in this function also work for buildText() and buildCSV()
def buildHTML(data, options):
    html = "<table border=1>"
    if 'a' in options['content']:   # -c command line option
        progLog.logWrite("Generating HTML output - ADDED.")
        html += "<tr><td colspan = {} align=center> <b>Entries Added</b></td></tr>\n".format(len(options['reportfields']))      # Title
        if len(data['added']['result']) == 0:  
            html += "<tr><td colspan = {} align=center>(None)</td></tr>".format(len(options['reportfields']))
        else:                                                                                                                   # Column headings
            html+= "<tr>"
            for i in range(len(data['added']['fieldnames'])):
                if data['added']['fieldnames'][i] in options['reportfields']:                                                   # Only add a column if it's specified in the .ini file  
                    html += "<th>{}</th>".format(data['added']['fieldnames'][i])
            html += "</tr>\n"
            for key, values in data['added']['result'].items():
                html += "<tr>"
                for i in range(len(data['added']['fieldnames'])):                                                               # Only add a column if it's specified in the .ini file 
                    if data['added']['fieldnames'][i] in options['reportfields']:
                        html += '<td>{}</td>'.format(values[data['added']['fieldnames'][i]])
                html += '</tr>\n'
   
    if 'r' in options['content']:    # -c command line option
        progLog.logWrite("Generating HTML output - REMOVED.")
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
        progLog.logWrite("Generating HTML output - SAME.")
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
    html += '<br>Records examined: {}<br>'.format(getRunIdCount(options['run_id']))
    html += '<br>Report generated by <a href="{}">arcomp</a> version {} ({})<br>'.format(gitSourceUrl, version[0], version[1])
    return html

def buildText(data, options):
    text = ''

    if 'a' in options['content']:
        progLog.logWrite("Generating Text output - ADDED.")
        text += "Entries Added\n"
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
        progLog.logWrite("Generating Text output - REMOVED.")
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
        progLog.logWrite("Generating Text output - SAME.")
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

    text += '\nRecords examined: {}\n'.format(getRunIdCount(options['run_id']))
    text += '\nReport generated by arcomp ({}) Version {} ({})\n'.format(gitSourceUrl, version[0], version[1])
    return text

def buildCSV(data, options):
    text=''
    
    if 'a' in options['content']:
        progLog.logWrite("Generating CSV output - ADDED.")
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
        progLog.logWrite("Generating CSV output - REMOVED.")
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
        progLog.logWrite("Generating CSV output - SAME.")
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

    text += '\nRecords examined: {}\n'.format(getRunIdCount(options['run_id']))
    text += '\nReport generated by arcomp ({}) Version {} ({})\n'.format(gitSourceUrl, version[0], version[1])
    return text

# Write the output report(s) to files, if specified on the command line
def writeFiles(data, options):
    progLog.logWrite("Writing reports to files.")
    for item in options['write'].items():
        if item[1].lower() == 'text':                   # Convert to text
            output = buildText(data, options)
        elif item[1].lower() == 'html':                 # Convert to HTML
            output = buildHTML(data, options)
        elif item[1].lower() == 'csv':                  # Convert to CSV
            output = buildCSV(data, options)
        else:                                   # Data is already in JSON format internally
            output = data

        progLog.logWrite("Writing output {} to {} file.".format(item[0], item[1].lower()))
        outfile = open('{}\\{}'.format(options['datapath'], item[0]),'w')
        if item[1].lower() in ['text','html','csv']:   # Write data to file    
            outfile.write(output)
        else:                                   # Json data requires special call
            json.dump(output, outfile)
        outfile.close()

# Send the report out via email
def sendEmail(data, options, inifile):
    progLog.logWrite("Sending email")
    try:
        serverconnect = smtplib.SMTP(options['email']['server'],options['email']['port'])
        if options['email']['encryption'] != None:   # Do we need to use SSL/TLS?
            try:
                tlsContext = ssl.create_default_context()
                serverconnect.starttls(context=tlsContext)
            except Exception as e:
                oops("TLS initiation errror")
        try:
            pw = iniFile.getIniOption('email','password')                               # Get password now so it's not stored in memory long-term
            retVal, retMsg = serverconnect.login(options['email']['account'], pw)  
        except:
            oops("Server login error")
    except (smtplib.SMTPAuthenticationError, smtplib.SMTPConnectError, smtplib.SMTPSenderRefused):
        e = sys.exc_info()[0]
        oops("Server authentication error")

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

# Send report data to syslog.
# Fields are prer-selected here, not based on the [fields] section of the .ini file
# See the documentation for an approproate GROK pattern to use with your syslog or SIEM system.
def sendSyslog(data, options):
    progLog.logWrite("Sending log to syslog server: [{}:{}]".format(options['syslog']['server'], options['syslog']['port']))
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

# Print the full arcomp run history, including run_ids and dates. Used to find a specific run_id to delete from the database with the -R option
def printHistory():
    progLog.logWrite("Printing run_id history")
    curs = db.execSqlStmt("SELECT DISTINCT run_id FROM history ORDER BY run_id ASC")
    runids = curs.fetchall()
    for i in range(len(runids)):
        id = runids[i][0]
        print("{}   ({}-{}-{}  {}:{}:{}.{})".format(id, id[0:4], id[4:6], id[6:8], id[9:11], id[11:13], id[13:15], id[16:]))
    return

def deleteRunID(runid):
    progLog.logWrite("Deleting run_id: [{}]".format(runid))
    curs = db.execSqlStmt("SELECT run_id FROM history WHERE run_id = '{}'".format(runid))
    result = curs.fetchall()
    if len(result) == 0:
        print("No such run_id: {}".format(runid))
        return None

    curs = db.execSqlStmt("DELETE FROM history WHERE run_id = '{}'".format(runid))
    result = curs.fetchall()
    return None

def getRunIdCount(runid):
    curs = db.execSqlStmt("SELECT COUNT(*) FROM history WHERE run_id='{}'".format(runid))
    count = curs.fetchone()[0]
    return count


##### Let's Go! #####
if __name__ == "__main__":

    # Global options dictionary holds all the program options, either default, from the .ini file, or from the commmand line
    options = {}
    options['progpath'] = os.path.dirname(os.path.realpath(sys.argv[0]))  # Get program home directory
    options['hostname'] = socket.gethostname()
    options['run_id'] = datetime.now().strftime("%Y%m%d-%H%M%S-%f") # Each run gets its own run identifier. All entries from the same run have the same run_id.
    options['version'] = version
    options['gitSourceUrl'] = gitSourceUrl
    options['copyright'] = copyright

    # Open and read the .ini file
    iniFile = IniOptions(options['progpath'] + '\\arcomp.ini')              # Class to handle .ini file operations
    options['autorunspath'] = iniFile.getIniOption('main','autorunspath')    # Path to autorunsc.exe. If Null, assume it's in thre Windows %PATH%
    options['datapath'] = iniFile.getIniOption('main','datapath')            # Path to data files. If Null, assume it's in the same directory as this program
    if options['datapath'] is None:                                         # Use default data path
        options['datapath'] = options['progpath']
    if options['datapath'][-1:] == '\\':
        options['datapath'] = options['datapath'][:-1]                      # Remove any trailing '\' since the rest of the program assumes it's not there
    options['reportfields'] = list({key: value for key, value in iniFile.getIniSection('fields').items() if value.lower() == 'true'})     # List of fields from .ini file [report] section to use in report output

    # Get signers and companies to ignore
    options['ignore_signer'] = {}
    if iniFile.hasSection('ignore_signer'):
        options['ignore_signer'] = iniFile.getIniSection('ignore_signer')
    options['ignore_company'] = {}
    if iniFile.hasSection('ignore_company'):
        options['ignore_company'] = iniFile.getIniSection('ignore_company')

    # Open log file
    progLog = Logger(options['datapath'] + '\\arcomp.log')
    progLog.logWrite('program path=[{}] run_id=[{}] version=[{}]'.format(options['progpath'], options['run_id'], options['version']))
    progLog.logWrite('autorunspath=[{}] datapath=[{}] reportfields=[{}]'.format(options['autorunspath'], options['datapath'], options['reportfields']))

    # Get and process command line arguments
    progArgs = processCmdLineArgs()
    options['file'] = progArgs.file

    if progArgs.write is not None:      # output files specified on the command line
        options['write'] = {}           # Dictionary of output files to write to
        for i in range(len(progArgs.write)):
            fname,type = progArgs.write[i].split(",")
            if type.lower() not in ['text','html','csv','json']:
                progLog.logWrite("Command line error, -w option. Invalid type: {}. Filetype must be 'text', 'html', 'csv', or 'json'".format(type))
                oops("Command line error, -w option. Invalid type: {}. Filetype must be 'text', 'html', 'csv', or 'json'".format(type))
            options['write'][fname] = type.lower()
    
    if progArgs.syslog is not None:     # Output to syslog is specified. Format is <server>[:port]
        options['syslog'] = {}
        syslogspec = progArgs.syslog.split(':')
        options['syslog']['server'] = syslogspec[0]
        options['syslog']['port'] = 514         # Syslog default port
        if len(syslogspec) == 2:                # Port specified
            options['syslog']['port'] = int(syslogspec[1])

    # Check if specifying 'added,' 'removed,' or 'same' sections in the report
    if progArgs.content is None:
        options['content'] = 'ars'
    else:
        for i in range(len(progArgs.content)):
            if progArgs.content[i] not in ['a','r','s']:
                progLog.logWrite("--content option: invalid option: '{}'. Must be a combination of 'a', 'r', and/or 's'".format(progArgs.content[i]))
                oops("--content option: invalid option: '{}'. Must be a combination of 'a', 'r', and/or 's'".format(progArgs.content[i]))
        options['content'] = progArgs.content

    # Check for email options
    options['email'] = iniFile.getIniSection('email')
    options['email']['send'] = progArgs.email
    options['email']['password'] = None                 # Do not store password until it's necessary to send email

    # Open and prep database
    db = Database(options['datapath'] + '\\arcompdata.db')
    db.dbSetup()
    options['dbfields'] = db.getTableFieldNames('history')      # Get names of the fields in the history table. This will come in handy later.
    
    # Need to just print history?
    if progArgs.runhistory is True:
        progLog.logWrite("Printing run history.")
        printHistory()
        db.dbCommit()
        db.dbClose()
        exit(0)

    # Need to delete a run_id?
    if progArgs.runremove is not None:
        progLog.logWrite("Removing run_id [{}].".format(progArgs.runremove))
        deleteRunID(progArgs.runremove)
        db.dbCommit()
        db.dbClose()
        exit(0)

    # Get last run_id. This will be used to compare against the current run_id.
    options['last_runid'] = getLastRunId()

    # Are we processing a command-line file or letting autorunsc.exe do its thing?
    if options['file'] is None:                 # There's no specific file to process. Execute autorunsc.exe and collect output file
        cmdline = '\"\"{}\" -a * -c -h -s -v -vt -o \"\"{}\\aroutput.csv\" -nobanner'.format(options['autorunspath'],options['datapath'])  
        progLog.logWrite("Running autoruns. Command line=[{}].".format(cmdline))
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

    # Close database and exit
    progLog.logWrite("Closing program.")
    db.dbCommit()
    db.dbClose()

