Attribute VB_Name = "Loggers"
Option Explicit
' =============================================================================
' Module:        Logger
' Author:        Mark Uildriks, codevba.com
' Description:   Logger functions to simplify use of Logger. Also includes demo & test procedures
' Office version 2016 and higher
' Dependencies:  None
' License:       MIT License
' Version        1.0
' Repository:    https://github.com/codevba-com/logger
' =============================================================================

Public Function MyLogger(Optional Restart As Boolean = False) As Logger
'Encapsulating the logger object in a function you can use the logger without worrying if it has been instantiated
'Using Static variable prevents the Logger variable from having to be intialized with each use
'Run 'MyLogger Restart:=True' in Immediate Window (Ctrl-G) to pick up changed properties
    Static sLogger As Logger 'value survives between calls
    If Restart Then Set sLogger = Nothing
    If sLogger Is Nothing Then
        Set sLogger = New Logger
        'easily switch where to log
        sLogger.SinkType = eSinkTypeImmediate
        'sLogger.SinkType = eSinkTypeFile
        'sLogger.SinkType = eSinkTypeAccessTable 'in MS Access only
        'Not in MS Access only
        'LogRecordFormat - default depends on SinkType
        'sLogger.LogRecordFormat = eLogRecordFormatCompact
        'sLogger.LogRecordFormat = eLogRecordFormatSimple
        'sLogger.LogRecordFormat = eLogRecordFormatJson
    End If
    Set MyLogger = sLogger
End Function

Public Function LoggerFile() As Logger 'short version of MyLogger - using defaults
    Static sLogger As Logger
    If sLogger Is Nothing Then Set sLogger = New Logger: sLogger.SinkType = eSinkTypeFile
    'sLogger.FilePath = "c:\temp\log.txt" 'default folder is next to the application document.
    Set LoggerFile = sLogger
End Function

Public Function LoggerAccessTable() As Logger ' Only available in MS Access
    Static sLogger As Logger
    If sLogger Is Nothing Then Set sLogger = New Logger: sLogger.SinkType = eSinkTypeAccessTable
    'sLogger.CreateLogTable 'default table tblLog created on first use
    Set LoggerAccessTable = sLogger
End Function

Public Function LoggerImmediateWindow() As Logger
    Static sLogger As Logger
    If sLogger Is Nothing Then Set sLogger = New Logger: sLogger.SinkType = eSinkTypeImmediate
    Set LoggerImmediateWindow = sLogger
End Function

Public Sub TestMyLogger()
    MyLogger.Log Message:="Test Logger File Start", Level:=eLogLevelInfo, Source:="LogTests.TestLoggerFile"
    MyLogger.Log Message:="Test Logger File End", Level:=eLogLevelInfo, Source:="LogTests.TestLoggerFile"
    MyLogger.Log Message:="Test Logger File End", Level:=eLogLevelInfo, Source:=""
End Sub

'----DEMO & TESTS-------------

'Public Sub TestLoggerDefault()
''If you do not specify a SinkType, it uses the default, see meDefaultSinkType
'    Set LoggerDefault = New Logger
'    LoggerDefault.Log "Test Logger Default"
'End Sub

Public Sub TestLoggerImmediateWindow()
    LoggerImmediateWindow.Log "Test Logger Immediate Window"
End Sub

Public Sub TestLoggerFile()
    With LoggerFile
        .SinkType = eSinkTypeFile
        .Log Message:="Test Logger File Start", Level:=eLogLevelInfo, Source:="LogTests.TestLoggerFile"
        .Log Message:="Test Logger File End", Level:=eLogLevelInfo, Source:="LogTests.TestLoggerFile"
    End With
End Sub

Public Sub TestLoggerAccessTable() 'can only be used in MS Access
    With LoggerAccessTable
        .SinkType = eSinkTypeAccessTable
        .Log "Test Logger Access Table", Level:=eLogLevelWarning, Source:="LogTests.TestLoggerAccessTable"
    End With
End Sub

