# WeBeeLogger - VBA Excel Add-In for logging
WeBeeLogger is a powerful and very flexible solution for logging.

You can use builtin handlers, processors and formatters or you can develop your owns - whatever suits you better.

The biggest advantage of this solution is that is fully OOP designed and developed (at least as much as VBA allows it ;P) - as there is no such a solutions on the market.

## Installation and basic usage
### Installation
To install this Add-In - just download newest binary .xlam file from this repository and add it to Excel via AddIns on Developer Tab.

To use it in your project add reference to this add-in and check below code examples.

### Examples
#### Create logger instance with default handlers, processors and formatters
```vb
Dim log As WeBeeLogger.LoggerInterface
Set log = WeBeeLogger.Factory.getImmediateLogger()

log.info "This is your first log message!"
```

Above will print out this message in immediate window:
```
[2020/03/21 17:10:00] NULL.info: "This is your first log message" [-]
```

#### Create logger named instance
It is a good practice to name your logger.
```vb
Dim log As WeBeeLogger.LoggerInterface
Set log = WeBeeLogger.Factory.getImmediateLogger("BEST_LOGGER")

log.info "This is your second log message!"
```

Above will print out this message in immediate window:
```
[2020/03/21 17:10:00] BEST_LOGGER.info: "This is your second log message" [-]
```

#### Automatically log error messages
```vb
Dim log As WeBeeLogger.LoggerInterface
Set log = WeBeeLogger.Factory.getImmediateLogger("BEST_LOGGER")

On Error GoTo EH

Err.Raise 1000, "Test code", "You use it wrongly"

EH:
log.error "Something went very wrong!"
```

Above will print out this message in immediate window:
```
[2020/03/21 17:10:00] BEST_LOGGER.error: "Something went very wrong! [Error: 1000: You use it wrongly (Test code)]" [-]
```

#### Add additional context data
```vb
Dim log As WeBeeLogger.LoggerInterface
Set log = WeBeeLogger.Factory.getImmediateLogger("BEST_LOGGER")

On Error GoTo EH

Err.Raise 1000, "Test code", "You use it wrongly"

EH:
Dim ctx As New Scripting.Dictionary

ctx("userName") = "adam.wojciechowski"

log.error "Something went very wrong!", ctx
```

Above will print out this message in immediate window:
```
[2020/03/21 17:10:00] BEST_LOGGER.error: "Something went very wrong! [Error: 1000: You use it wrongly (Test code)]" [userName -> 'adam.wojciechowski']
```

#### Add additional context data and message placeholders
```vb
Dim log As WeBeeLogger.LoggerInterface
Set log = WeBeeLogger.Factory.getImmediateLogger("BEST_LOGGER")

On Error GoTo EH

Err.Raise 1000, "Test code", "You use it wrongly"

EH:
Dim ctx As New Scripting.Dictionary

ctx("userName") = "adam.wojciechowski"
ctx("parsedFile") = "C:\test.txt"

log.error "No access to {parsedFile}!", ctx
```

Above will print out this message in immediate window:
```
[2020/03/21 17:10:00] BEST_LOGGER.error: "No access to C:\test.txt! [Error: 1000: You use it wrongly (Test code)]" [userName -> 'adam.wojciechowski', parsedFile -> 'C:\test.txt']
```

#### Send logs to immediate window and to file
```vb
Dim log As WeBeeLogger.LoggerInterface
Dim cLog As WeBeeLogger.ConstructableInterface

Set cLog = WeBeeLogger.Factory.getFileLogger(loggerName:="BEST_LOGGER")
cLog.registerHandler WeBeeLogger.Factory.getImmediateHandler()

Set log = cLog

On Error GoTo EH

Err.Raise 1000, "Test code", "You use it wrongly"

EH:
Dim ctx As New Scripting.Dictionary

ctx("userName") = "adam.wojciechowski"
ctx("parsedFile") = "C:\test.txt"

log.error "No access to {parsedFile}!", ctx
```

Above will print out this message in immediate window and in the same will store it in file '%USERPROFILE%\VBA_Logs\20200321_vba.log':
```
[2020/03/21 17:10:00] BEST_LOGGER.error: "No access to C:\test.txt! [Error: 1000: You use it wrongly (Test code)]" [userName -> 'adam.wojciechowski', parsedFile -> 'C:\test.txt']
```

#### Pass logger to your classes

```vb
' class "MyProject.XYZ" code:
Option Explicit

Implements WeBeeLogger.LoggerAwareInterface

Private log As WeBeeLogger.LoggerInterface

Private Function LoggerAwareInterface_setLogger(ByRef logger As WeBeeLogger.LoggerInterface)

    Set log = logger

End Function

Public Sub someThing()

    log.info "Logger is here!"

End Sub

' class usage:
Dim c As New MyProject.XYZ
Dim cLog As WeBeeLogger.LoggerAwareInterface

Set cLog = c

cLog.setLogger WeBeeLogger.Factory.getImmediateLogger(loggerName:="BEST_LOGGER")

c.someThing
```

Above will print out this message in immediate window:
```
[2020/03/21 17:10:00] BEST_LOGGER.info: "Logger is here!" [-]
```
