# vba-logger  
A versatile logging class for VBA with multiple sinks and formatters.

## Logger Class

The `Logger` class stores structured log data using the `LogRecord` type:

| Field         | Description |
|---------------|-------------|
| **Timestamp** | System date/time (automatic) |
| **Level**     | Enumerated severity level |
| **LevelName** | Readable label (e.g., `INFO`, `ERROR`) |
| **Message**   | Event details |
| **Source**    | Event source identifier |
| **UserName**  | `Environ$("USERNAME")` |
| **MachineName** | `Environ$("COMPUTERNAME")` |

### Supported Sinks
You can route log output to any of these destinations:

1. **File**  
   - Set `FilePath` or use default: *DocumentsFolder/log.txt*
2. **MS Access Table**  
   - Set `LogTableName` or use default: *tblLog*
3. **Immediate Window**

Use the `SinkType` property to select the destination.  
For multiple sinks, simply create multiple `Logger` instances.

### Record Formatters
Available formatters:

- **Compact** — default for Immediate Window  
- **Simple** — default for File  
- **Json**

Select a formatter using the `LogRecordFormat` property.

At the top of the class module you’ll find six constants that define default behavior.

---

## Loggers Module

The `Loggers` module provides three ready‑to‑use logger functions:

- `LoggerFile`
- `LoggerAccessTable`
- `LoggerImmediateWindow`

And a customizable one:

- `MyLogger` — lets you define both `SinkType` and `LogRecordFormat`.

### Example Usage

```vba
MyLogger.Log _
    Message:="Test Logger File Start", _
    Level:=eLogLevelInfo, _
    Source:="LogTests.TestLoggerFile"
```
```text
2026-03-16 12:46:52 [INFO]  LogTests.TestLoggerFile – Test Logger File Start
2026-03-16 12:46:52 [DEBUG] LogTests.TestLoggerFile – Running step 1
2026-03-16 12:46:53 [WARN]  LogTests.TestLoggerFile – Unexpected value, using default
2026-03-16 12:46:54 [ERROR] LogTests.TestLoggerFile – Failed to open file
2026-03-16 12:46:55 [INFO]  LogTests.TestLoggerFile – Test Logger File End


### Notes
Inside `MyLogger` you’ll find an explanation of how a `Static` variable ensures the logger is instantiated only once.

The module also includes a **DEMO & TESTS** section so you can quickly verify functionality.


