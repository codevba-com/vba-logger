VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' =============================================================================
' Module:        Logger
' Author:        Mark Uildriks, codevba.com
' Description:   Logger for VBA projects.
' Office version 2016 and higher
' Dependencies:  None
' License:       MIT License
' Version        1.0
' Repository:    https://github.com/codevba-com/logger
' =============================================================================

Private Const mcstrDefaultDocumentFolderAlternative As String = "c:\temp"
Private Const mcstrDefaultLogFileName As String = "log.txt"
Private Const mcstrDefaultLogTableName As String = "tblLog"
Private Const meDefaultSinkType As Long = 1 'eSinkTypeFile
Private Const meDefaultLogRecordFormatSinkFile As Long = 1 'eLogRecordFormatSimple: all record fields
Private Const meDefaultLogRecordFormatSinkImmediateWindow As Long = 2 'eLogRecordFormatCompact: LevelName:Message

Private LogTableIsPresent As Boolean

Private meSinkType As eSinkType
Private meLogRecordFormat As eLogRecordFormat
Private mstrFilePath As String
Private mstrLogTableName As String

Private Type LogRecord
    Timestamp As Date
    Level As Long
    LevelName As String
    Message As String
    Source As String
    UserName As String
    MachineName As String
End Type

Public Enum eSinkType 'Logger Sink Type - when setting up a new logger
    eSinkTypeFile = 1
    eSinkTypeImmediate = 2
    eSinkTypeAccessTable = 3
End Enum

Public Enum eLogLevel
    eLogLevelDebug = 10
    eLogLevelInfo = 20
    eLogLevelWarning = 30
    eLogLevelError = 40
    eLogLevelCritical = 50
End Enum

Public Enum eLogRecordFormat
    eLogRecordFormatNotSpecified = 0
    eLogRecordFormatSimple = 1
    eLogRecordFormatCompact = 2
    eLogRecordFormatJson = 3
End Enum

Private Const mclngErrNumberBase As Long = vbObjectError + 4941
Private Const mclngErrNumberMicrosoftAccessOnly As Long = mclngErrNumberBase + 1

Public Sub Log(Message As String, Optional Level As eLogLevel = eLogLevelDebug, Optional Source As String = "")
    Dim rec As LogRecord
    rec.Timestamp = Now
    rec.Level = Level
    rec.LevelName = LevelName(Level)
    rec.Message = Message
    rec.Source = Source
    rec.UserName = Environ$("USERNAME")
    rec.MachineName = Environ$("COMPUTERNAME")
    Write_ rec
End Sub

Private Function LevelName(LogLevel As eLogLevel) As String
    Select Case LogLevel
    Case eLogLevelDebug: LevelName = "DEBUG"
    Case eLogLevelInfo: LevelName = "INFO"
    Case eLogLevelWarning: LevelName = "WARNING"
    Case eLogLevelError: LevelName = "ERROR"
    Case eLogLevelCritical: LevelName = "CRITICAL"
    End Select
End Function

'---LOGSINK-------
Private Sub Write_(Record As LogRecord)
    If meSinkType = 0 Then meSinkType = meDefaultSinkType
    Select Case SinkType
    Case eSinkTypeFile
        WriteToFile Record
    Case eSinkTypeImmediate
        WriteToImmediateWindow Record
    Case eSinkTypeAccessTable
        WriteToAccessTable Record
    End Select
End Sub

Private Function FormatRecord(ByRef Record As LogRecord) As String
'Extend with your own Format functions if required
    Select Case LogRecordFormat
    Case eLogRecordFormatCompact
        FormatRecord = LogRecordFormatCompact(Record:=Record)
    Case eLogRecordFormatSimple
        FormatRecord = LogRecordFormatSimple(Record:=Record)
    Case eLogRecordFormatJson
        FormatRecord = LogRecordFormatJson(Record:=Record)
    End Select
End Function

Private Sub WriteToFile(Record As LogRecord)
    Dim text As String
    text = FormatRecord(Record)
    
    Dim f As Integer: f = FreeFile
    Open FilePath For Append As #f
    Print #f, text
    Close #f
End Sub

Private Sub WriteToImmediateWindow(Record As LogRecord)
    Debug.Print FormatRecord(Record)
End Sub

Private Sub WriteToAccessTable(Record As LogRecord) 'MS Access only
    Const dbOpenDynaset As Long = 2
    If Not LogTableIsPresent Then
        CreateLogTable
    End If
    
    Dim app As Object
    Set app = Application
    Dim db As Object 'DAO.Database
    Set db = app.CurrentDb
    Dim rst As Object 'DAO.Recordset
    Set rst = db.OpenRecordset("tblLog", dbOpenDynaset)
    With rst
        .AddNew
        ![Timestamp] = Record.Timestamp
        ![Level] = Record.Level
        ![LevelName] = Record.LevelName
        ![Source] = Record.Source
        ![Message] = Record.Message
        ![UserName] = Record.UserName
        ![MachineName] = Record.MachineName
        .Update
        .Close
    End With
    Set rst = Nothing
    Set db = Nothing
    Set app = Nothing
End Sub

Public Property Get SinkType() As eSinkType
    SinkType = meSinkType
End Property

Public Property Let SinkType(SinkType As eSinkType)
' you can change or extend LogRecordFormat used as you require
    meSinkType = SinkType
    If LogRecordFormat = eLogRecordFormatNotSpecified Then
        Select Case SinkType
        Case eSinkTypeFile
            LogRecordFormat = meDefaultLogRecordFormatSinkFile
        Case eSinkTypeImmediate
            LogRecordFormat = meDefaultLogRecordFormatSinkImmediateWindow
        Case eSinkTypeAccessTable 'does not use LogRecordFormat
            OnlyForMicrosoftAccess
        End Select
    End If
End Property

Public Property Get LogRecordFormat() As eLogRecordFormat
    LogRecordFormat = meLogRecordFormat
End Property

Public Property Let LogRecordFormat(LogRecordFormat As eLogRecordFormat)
' you can change or extend LogRecordFormat used as you require
    meLogRecordFormat = LogRecordFormat
End Property

Private Function OnlyForMicrosoftAccess() As Boolean
'test precondition for procedures that only work in MS Access
    If Application.Name <> "Microsoft Access" Then Err.Raise mclngErrNumberMicrosoftAccessOnly, "LogSink.SinkType", "Only works in Microsoft Access"
End Function

Private Function DocumentFolder() As String
'used for WriteToFile if FilePath has not been set s default folder.
'uses the folder where the application document (file) is located.
'when working on unsaved document it uses mcstrDefaultDocumentFolderAlternative (here: "c:\temp")
'returns folder without \
    Dim objApplication As Object: Set objApplication = Application
    On Error Resume Next
    Select Case Right$(Application.Name, Len(Application.Name) - 10)
    Case "Access"
        DocumentFolder = objApplication.CurrentProject.Path
    Case "Excel"
        DocumentFolder = objApplication.ActiveWorkbook.Path
    Case "Word"
        DocumentFolder = objApplication.ActiveDocument.Path
    Case "PowerPoint"
        DocumentFolder = objApplication.ActivePresentation.Path
    End Select
    If (Len(DocumentFolder) = 0) Then DocumentFolder = mcstrDefaultDocumentFolderAlternative
End Function
Public Property Get FilePath() As String
    If Len(mstrFilePath) > 0 Then
        FilePath = mstrFilePath
    Else
        FilePath = DocumentFolder & "\" & mcstrDefaultLogFileName
    End If
End Property
Public Property Let FilePath(rData As String)
    mstrFilePath = rData
End Property

Private Property Get LogTableName() As String
    OnlyForMicrosoftAccess
    If Len(mstrLogTableName) > 0 Then
        LogTableName = mstrLogTableName
    Else
        LogTableName = mcstrDefaultLogTableName
    End If
End Property
Public Property Let LogTableName(rData As String)
    OnlyForMicrosoftAccess
    mstrLogTableName = rData
End Property

'------lOGRECORD------
Private Function LogRecordToString(Record As LogRecord, LogRecordFormat As eLogRecordFormat) As String
    Select Case LogRecordFormat
    Case eLogRecordFormatSimple: LogRecordToString = LogRecordFormatSimple(Record)
    Case eLogRecordFormatCompact: LogRecordToString = LogRecordFormatCompact(Record)
    Case eLogRecordFormatJson: LogRecordToString = LogRecordFormatJson(Record)
    End Select
End Function

' Simple, readable line format
Private Function LogRecordFormatSimple(Record As LogRecord) As String
    LogRecordFormatSimple = _
    Format(Record.Timestamp, "yyyy-mm-dd hh:nn:ss") & _
    " [" & Record.LevelName & "] " & _
    Record.Source & " – " & Record.Message
End Function

' Compact format for Immediate Window
Private Function LogRecordFormatCompact(Record As LogRecord) As String
    LogRecordFormatCompact = Record.LevelName & ": " & Record.Message
End Function

' JSON format for structured logs
Private Function LogRecordFormatJson(Record As LogRecord) As String
    Dim s As String
    s = "{""ts"":""" & Format(Record.Timestamp, "yyyy-mm-dd\THH:NN:SS") & """," & _
    """level"":""" & EscapeJson(Record.LevelName) & """," & _
    """source"":""" & EscapeJson(Record.Source) & """," & _
    """msg"":""" & EscapeJson(Record.Message) & """," & _
    """user"":""" & EscapeJson(Record.UserName) & """," & _
    """machine"":""" & EscapeJson(Record.MachineName) & """}"
    LogRecordFormatJson = s
End Function

Private Function EscapeJson(s As String) As String
    EscapeJson = Replace(s, """", "\""")
End Function

'-------------------------
Public Sub CreateLogTable(Optional ByVal TableName As String = "tblLog")
    OnlyForMicrosoftAccess
    Const dbLong As Long = 4
    Const dbDate As Long = 8
    Const dbText As Long = 10
    Const dbMemo As Long = 12
    Const dbAutoIncrField As Long = 16
    Dim app As Object
    Set app = Application
    Dim db As Object 'DAO.Database
    Dim tdf As Object 'DAO.TableDef
    Set db = app.CurrentDb
    
    'exit if table exists
    For Each tdf In db.TableDefs
        If StrComp(tdf.Name, TableName, vbTextCompare) = 0 Then 'string1 is equal to string2
            LogTableIsPresent = True
            Exit Sub
        End If
    Next
    
    Set tdf = db.CreateTableDef(TableName)
    Dim fld As Object 'DAO.Field
    Set fld = tdf.CreateField("ID", dbLong)
    fld.Attributes = fld.Attributes Or dbAutoIncrField
    tdf.Fields.Append fld
    tdf.Fields.Append tdf.CreateField("Timestamp", dbDate)
    tdf.Fields.Append tdf.CreateField("Level", dbLong)
    tdf.Fields.Append tdf.CreateField("LevelName", dbText, 20)
    tdf.Fields.Append tdf.CreateField("Source", dbText, 255)
    tdf.Fields.Append tdf.CreateField("Message", dbMemo)
    tdf.Fields.Append tdf.CreateField("UserName", dbText, 50)
    tdf.Fields.Append tdf.CreateField("MachineName", dbText, 50)
    db.TableDefs.Append tdf
    Set tdf = db.CreateTableDef(TableName)
    app.RefreshDatabaseWindow
    LogTableIsPresent = True
    Set fld = Nothing
    Set tdf = Nothing
    Set db = Nothing
    Set app = Nothing
End Sub


