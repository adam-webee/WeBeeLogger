Attribute VB_Name = "Factory"
''
' WeBeeLogger - VBA Logger Add-In for Excel
' Copyright (C) 2020  Adam Wojciechowski
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <https://www.gnu.org/licenses/>.
''
Option Explicit

''
' Builds BeeLogger instance that will put logs to immediate window.
'
' Will use placeholders, error and context processors for processing log record.
'
' @param String                         loggerName [opt] name for the logger, by default empty
' @param WeBeeLogger.FormatterInterface formatter  [opt] if not provided @see: WeBeeLogger.LineFormatter will be used
'
' @return WeBeeLogger.LoggerInterface
''
Public Function getImmediateLogger( _
    Optional ByVal loggerName As String = vbNullString, _
    Optional ByRef formatter As WeBeeLogger.FormatterInterface = Nothing _
) As WeBeeLogger.LoggerInterface

    Dim l As WeBeeLogger.ConstructableInterface
    Set l = New WeBeeLogger.BeeLogger

    With l
        Set getImmediateLogger = .construct(loggerName)
        .registerProcessor getPlaceholdersProcessor()
        .registerProcessor getErrorProcessor()
        .registerProcessor getContextProcessor()
        .registerHandler getImmediateHandler(formatter)
    End With

End Function

''
' Builds BeeLogger instance that will put logs to file.
'
' Will use placeholders, error and context processors for processing log record.
'
' @param String                         loggerName [opt] name for the logger, by default empty
' @param String                         filePath   [opt] path to log file, by default %USERPROFILE%\VBA_Logs\yyyymmdd_vba.log
' @param WeBeeLogger.FormatterInterface formatter  [opt] if not provided @see: WeBeeLogger.LineFormatter will be used
'
' @return WeBeeLogger.LoggerInterface
''
Public Function getFileLogger( _
    Optional ByVal loggerName As String = vbNullString, _
    Optional ByVal filePath As String = vbNullString, _
    Optional ByRef formatter As WeBeeLogger.FormatterInterface = Nothing _
) As WeBeeLogger.LoggerInterface

    Dim l As WeBeeLogger.ConstructableInterface
    Set l = New WeBeeLogger.BeeLogger

    With l
        Set getFileLogger = .construct(loggerName)
        .registerProcessor getPlaceholdersProcessor()
        .registerProcessor getErrorProcessor()
        .registerProcessor getContextProcessor()
        .registerHandler getFileHandler(filePath, formatter)
    End With

End Function

''
' Builds NullLogger instance to be used as fulfillment of dependency in places that do not need to log anything.
'
' @param String loggerName [opt] name for the logger, by default empty
'
' @return WeBeeLogger.LoggerInterface
''
Public Function getNullLogger( _
    Optional ByVal loggerName As String = vbNullString _
) As WeBeeLogger.LoggerInterface

    Dim l As WeBeeLogger.ConstructableInterface
    Set l = New WeBeeLogger.NullLogger

    With l
        Set getNullLogger = .construct(loggerName)
    End With

End Function

''
' Builds processor that can parse placeholders in log messages.
'
' @return WeBeeLogger.PlaceholdersProcessor
''
Public Function getPlaceholdersProcessor() As WeBeeLogger.PlaceholdersProcessor

    Set getPlaceholdersProcessor = New WeBeeLogger.PlaceholdersProcessor

End Function

''
' Builds processor that can parse thrown error message and add it to log message.
'
' @return WeBeeLogger.ErrorProcessor
''
Public Function getErrorProcessor() As WeBeeLogger.ErrorProcessor

    Set getErrorProcessor = New WeBeeLogger.ErrorProcessor

End Function

''
' Builds processor that can change context structure in to string representation and append it to log message.
'
' @return WeBeeLogger.ContextProcessor
''
Public Function getContextProcessor() As WeBeeLogger.ContextProcessor

    Set getContextProcessor = New WeBeeLogger.ContextProcessor

End Function

''
' Builds formatter that format log record in to the one line string.
'
' @return WeBeeLogger.LineFormatter
''
Public Function getLineFormatter() As WeBeeLogger.LineFormatter

    Set getLineFormatter = New WeBeeLogger.LineFormatter

End Function

''
' Builds handler that puts log messages to immediate window.
'
' @param WeBeeLogger.FormatterInterface formatter [opt] formatter to be used by handler, by default @see: WeBeeLogger.LineFormatter
'
' @return WeBeeLogger.ImmediateHandler
''
Public Function getImmediateHandler( _
    Optional ByRef formatter As WeBeeLogger.FormatterInterface = Nothing _
) As WeBeeLogger.ImmediateHandler

    Dim h As WeBeeLogger.HandlerInterface
    Set h = New WeBeeLogger.ImmediateHandler

    If (Nothing Is formatter) Then
        Set h.formatter = getLineFormatter()
    Else
        Set h.formatter = formatter
    End If

    Set getImmediateHandler = h

End Function

''
' Builds handler that puts log messages to file.
'
' @param String                         filePath  [opt] path to file where to store log messages, default @see: WeBeeLogger.FileHandler.logFilePath
' @param WeBeeLogger.FormatterInterface formatter [opt] formatter to be used by handler, by default @see: WeBeeLogger.LineFormatter
'
' @return WeBeeLogger.ImmediateHandler
''
Public Function getFileHandler( _
    Optional ByVal filePath As String = vbNullString, _
    Optional ByRef formatter As WeBeeLogger.FormatterInterface = Nothing _
) As WeBeeLogger.FileHandler

    Dim h As WeBeeLogger.HandlerInterface

    With New WeBeeLogger.FileHandler
        Set h = .construct(filePath)
    End With

    If (Nothing Is formatter) Then
        Set h.formatter = getLineFormatter()
    Else
        Set h.formatter = formatter
    End If

    Set getFileHandler = h

End Function
