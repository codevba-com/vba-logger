# vba-logger - a versatile Logging class with multiple sinks and formatters.

##Logger class

With the Logger class you can store the following data (using Type LogRecord)
Timestamp    system date (automatic)
Level        enumerated values allowing sort based on severity
LevelName    readable, e.g. INFO or ERROR
Message      details of the event
Source       event source
UserName     Environ$("USERNAME")
MachineName  Environ$("COMPUTERNAME")

The Logger class supports three sinks:
1. File: specify FilePath - or keep default: DocumentFolder (application document + log.txt) 
2. MS Access table: set LogTableName or keep default name tblLog
3. Immediate Window
You can specify your log data destination (sink) setting property SinkType 
If you need multiple sinks, use multiple instances of the Logger class

For File and Immediate Window you use prefab log record formatters:
1. Compact - default for Immediate Window
2. Simple - default for File
3. Json
You can alter the formatter used by setting property LogRecordFormat.

At the top of the Logger class you can find 6 constants that determine the default behaviour. 

##Loggers module

The Loggers module contains three prefab Logger procedures: LoggerFile, LoggerAccessTable and LoggerImmediateWindow.
The fourth, MyLogger, let's you define one with SinkType and LogRecordFormat to tailor to your needs.
Inside the MyLogger procedure you find a small explanation on the use of static variable to have the Logger class only instantiate once. 


The Loggers module also contains a small DEMO & TESTS section. Here you can eaily verify if it will work for you.


