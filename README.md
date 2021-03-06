# arcomp
Arcomp is a Python script that is used in conjunction with the [Autoruns utility](https://docs.microsoft.com/en-us/sysinternals/downloads/Autoruns) from the Microsoft [Sysinternals package](https://docs.microsoft.com/en-us/sysinternals/). Arcomp compares the results of successive Autoruns executions and reports on what's been added, removed, and unchanged between runs. Some of arcomp's features include:

- Automatically collating and reporting on changes between Autoruns executions
- Send the report to a Text, HTML, CSV, or JSON file
- Email the report to an administrator
- Send the report to a syslog server or SIEM for further analysis

# Package Files

- arcomp.py - The Python script that performs the Autoruns analysis
- arcomp.ini.EXAMPLE - An example initialization file that provides runtime information for arcomp. See the *arcomp.ini File* section below for more details.
- arclaunch.bat.EXAMPLE - An example Windows batch script that executes arcomp.py with appropriate parameters. This can be useful if setting up arcomp under the Windows Task Scheduler, so that you can instruct Task Scheduler to simply execute the batch script rather than coding the arcomp parameters into the Task Scheduler options.

# Installing arcomp

To install arcomp, do the following:

1. Download and install Python 3.0 (or higher) on your Windows computer. Instructions are [here](https://www.python.org/downloads/).
2. Download and install Autoruns on your Windows computer. Instructions are [here](https://docs.microsoft.com/en-us/sysinternals/downloads/autoruns).
3. Clone the [arcomp GitHub repository](https://github.com/HandyGuySoftware/arcomp) or download the files to a directory on your local Windows machine.
4. Copy the arcomp.ini.EXAMPLE file to arcomp.ini and edit the options as appropriate
5. (Optional) Copy the arclaunch.bat.EXAMPLE file to arclaunch.bat and edit the variables at the top of the file as appropriate
6. Either:
   1. Run arcomp.py using the desired command line options, **or**
   2. Use arclaunch.bat to run arcomp.py.

# Usage

**C:\>** arcomp [-f \<filename>] [-w \<write-file>,\<type>] [-e] [-s \<syslog_server>[:\<port]] [-c \<a|r|s>] [-r] [-R \<run_id>]

| Option                        | Description                                                  |
| ----------------------------- | ------------------------------------------------------------ |
| -c \<a\|r\|s>                 | Specify the sections of the data to send in the report. Arcomp analyzes what information has been added ('a'), removed ('r'), or stayed the same ('s') between Autoruns executions. The resulting report will only include the sections specified by the -r option. By default, all sections are included in the report. However, since the majority of Autoruns entries do not change between executions, most users select only the 'a' and 'r' entries to see only what's been added or removed.<br /><br />Note: The -c option only affects the output for the Text, HTML, and CSV outputs from arcomp. The JSON and syslog outputs always contain the full data ('a', 'r', and 's'). |
| -e                            | Send the report via email. Email parameters are specified in the [email] section of the arcomp.ini file. |
| -f \<filename>                | Use \<filename> as the data input to the program. If the -f option is not used, the program will execute Autoruns and use the output of that run as input to arcomp. \<Filename> must be in Comma-Separated Value (CSV) format and must be created by Autoruns using the following command line options:<br /><br /> `'autorunsc.exe -a * -c -h -s -v -vt -o \<filename.csv> -nobanner'` |
| -r                            | Outputs the full arcomp run history, including the Run ID and the data/time the run was executed. The Run ID can be used to remove a run from the database using the -R option. |
| -R                            | Remove a run from the arcomp database. \<run_id> is the Run ID to remove. All entries in the database for that run_id will be deleted. |
| -s \<syslog_server>[:\<port>] | Send the report to a syslog or SIEM server. \<syslog_server> is the IP address or fully-qualified domain name of the server. [:\<port>] may be specified if the syslog server uses a non-standard port. If [:\<port>] is not specified, the default port is 514. |
| -w \<write-file>,\<type>      | Write the output report to a Text, HTML, CSV, or JSON file. <br />\<writefile> is the name of the file where the report will be written. <br /><br />\<type> is one of 'html', 'text', 'csv', or 'json'<br /><br />The -f option can be specified multiple times to create more than one format of output report. For example:<br /><br />`arcomp.py -w output.txt,text -w output.html,html -w output.json,json` |

# The arcomp.ini file

Arcomp.ini contains parameters that arcomp uses to run the program properly. Arcomp.ini **must** be located in the same directory as the arcomp.py file.

## [main] section

| Option        | Description                                                  | Example                                      |
| ------------- | ------------------------------------------------------------ | -------------------------------------------- |
| Autorunspath= | This is the full path to the autorunsc.exe program on your drive | C:\Users\me\Documents\Autoruns\Autorunsc.exe |
| datapath=     | This is the directory where arcomp keeps its data and output files. If any report files are generated using the -w option, tho reports will be created in this directory. If datapath is not specified, the program will use the directory where arcomp.py is located by default. | C:\Users\me\Documents\arcomp                 |

## [email] section

The [email] section is only used if the program is run with the -e option. Otherwise, it is ignored.

| Option      | Description                                                  | Example                                 |
| ----------- | ------------------------------------------------------------ | --------------------------------------- |
| server=     | This is the SMTP server that arcomp will use to send email   | smtp.gmail.com                          |
| port=       | The SMTP port to use on the server                           | 587                                     |
| encryption= | True/False option to instruct arcomp to use TLS with the SMTP server | True                                    |
| account=    | The account to use for logging into the SMTP server          | myaccount@gmail.com                     |
| password=   | The password to use for the SMTP server                      | mypassword                              |
| sender=     | The email address for the sender. Typically this is the same as your user account. | myaccount@gmail.com                     |
| sendername= | The 'friendly' name of the sender's email account            | Arcomp Reporter                         |
| receiver=   | The email address where the report will be sent              | receiver@companymail.com                |
| subject=    | The subject for the outgoing email                           | Autoruns Comparison Report - Systemname |

## [fields] section

The [fields] section indicates what fields to include in the text, HTML, and CSV reports. This section has no effect on the JSON or syslog outputs. To include a field on the report, set the entry for that field to 'True'. To leave a field out of the report, set the entry to 'False' or leave it blank.

## [ignore_signer] section

This section should contain a list of verified signers that can safely be ignored in the reporting results. The lines in this field are case sensitive and must match *exactly* the data as it appears in Autoruns. An example of this section is:

`[ignore_signer]`

`(Verified) Microsoft Windows`

`(Verified) Microsoft Corporation`

## [ignore_company] section

This section should contain a list of company names that can safely be ignored in the reporting results. The lines in this field are case sensitive and must match *exactly* the data as it appears in Autoruns. An example of this section is:

`[ignore_company]`

`Adobe Systems, Incorporated`

`Adobe Systems Incorporated`

`Microsoft Corporation`



Note: if an data line has *either* an ignored signer *or* an ignored company, the line will be ignored.

# Syslog Parsing

Arcomp can send output to a syslog or SIEM server using the -s option. The following Grok string can be used to parse the arcomp feed:

`"\<%{NUMBER:UNWANTED}\>\[%{DATA:runtime}\]\[%{DATA:hostname}\]\[%{DATA:loglevel}\]\[%{DATA:run_id}\]\[%{DATA:action}\](?<location>[^|]*)\|(?<entry>[^|]*)\|(?<description>[^|]*)\|(?<signer>[^|]*)\|(?<company>[^|]*)\|(?<imagepath>[^|]*)\|%{GREEDYDATA:launchstring}"`

This GROK string has been tested and used on an Elastic (ELK) stack. Some modification may be needed for other syslog or SIEM implementations. If you successfully create a parsing string for another platform, please let the developer know and this documentation will be updated.

# Known Issues and Limitations

- Error handling and logging is limited and sketchy. This will improve as the code matures.
- The program puts out a high volume of syslog messages in a short period of time. This may overwhelm slower syslog/SIEM servers and messages may get dropped as a result. Some sort of rate limiting function may be added in the future to address this.

# Distribution, Support, and Feedback

Arcomp is an open source utility distributed under the MIT Open Source License. 

Feedback, bug reports, and feature suggestions are welcome. To start a discussion, please open an Issue on the [Arcomp GitHub Issues Page](https://github.com/HandyGuySoftware/arcomp/issues). **Please do not** submit a pull request without first opening an Issue and discussing the update with the developer.

