VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BeeLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Implements WeBeeLogger.LoggerInterface
Implements WeBeeLogger.ConstructableInterface

''
' @var Scripting.Dictionary properties [
'   @var String                           name logger name
'   @var WeBeeLogger.HandlerInterface[]   handlers
'   @var WeBeeLogger.ProcessorInterface[] processors
' ]
''
Private properties As New Scripting.Dictionary

''
' @inheritdoc
''
Private Sub Class_Initialize()

    With properties
        .Item("name") = vbNullString
        Set .Item("processors") = New VBA.Collection
        Set .Item("handlers") = New VBA.Collection
    End With

End Sub

''
' @inheritdoc
''
Private Function ConstructableInterface_construct(Optional ByVal name As String = vbNullString) As WeBeeLogger.LoggerInterface

    properties.Item("name") = name
    Set ConstructableInterface_construct = Me

End Function

''
' @inheritdoc
''
Private Sub ConstructableInterface_registerHandler(ByRef value As WeBeeLogger.HandlerInterface)

    properties.Item("handlers").Add value

End Sub

''
' @inheritdoc
''
Private Sub ConstructableInterface_registerProcessor(ByRef value As WeBeeLogger.ProcessorInterface)

    properties.Item("processors").Add value

End Sub

''
' Process log record with all registered processors.
'
' @param WeBeeLogger.LogRecordInterface lr
'
' @return void
''
Private Sub processLogRecord(ByRef lr As WeBeeLogger.LogRecordInterface)

    Dim tempProcessor As WeBeeLogger.ProcessorInterface

    For Each tempProcessor In properties.Item("processors")
        tempProcessor.process lr
    Next tempProcessor

End Sub

''
' Handle log record with all registered handlers.
'
' @param WeBeeLogger.LogRecordInterface lr
'
' @return void
''
Private Sub handleLogRecord(ByRef lr As WeBeeLogger.LogRecordInterface)

    Dim tempHandler As WeBeeLogger.HandlerInterface

    For Each tempHandler In properties.Item("handlers")
        If (tempHandler.isHandling(lr)) Then
            If (tempHandler.handle(lr)) Then Exit Sub
        End If
    Next tempHandler

End Sub

''
' @inheritdoc
''
Private Function LoggerInterface_log(ByVal level As WeBeeLogger.LOG_LEVELS, ByVal message As String, Optional ByRef context As Scripting.IDictionary = Nothing)

    Dim lr As WeBeeLogger.LogRecordInterface

    With New WeBeeLogger.LogRecord
        Set lr = .construct(message, level, properties.Item("name"), context)
    End With

    processLogRecord lr
    handleLogRecord lr

End Function

''
' @inheritdoc
''
Private Function LoggerInterface_emergency(ByVal message As String, Optional ByRef context As Scripting.IDictionary = Nothing)

    LoggerInterface_log WeBeeLogger.LOG_LEVELS.levelEmergency, message, context

End Function

''
' @inheritdoc
''
Private Function LoggerInterface_alert(ByVal message As String, Optional ByRef context As Scripting.IDictionary = Nothing)

    LoggerInterface_log WeBeeLogger.LOG_LEVELS.levelAlert, message, context

End Function

''
' @inheritdoc
''
Private Function LoggerInterface_critical(ByVal message As String, Optional ByRef context As Scripting.IDictionary = Nothing)

    LoggerInterface_log WeBeeLogger.LOG_LEVELS.levleCritical, message, context

End Function

''
' @inheritdoc
''
Private Function LoggerInterface_error(ByVal message As String, Optional ByRef context As Scripting.IDictionary = Nothing)

    LoggerInterface_log WeBeeLogger.LOG_LEVELS.levelError, message, context

End Function

''
' @inheritdoc
''
Private Function LoggerInterface_warning(ByVal message As String, Optional ByRef context As Scripting.IDictionary = Nothing)

    LoggerInterface_log WeBeeLogger.LOG_LEVELS.levelWarning, message, context

End Function

''
' @inheritdoc
''
Private Function LoggerInterface_notice(ByVal message As String, Optional ByRef context As Scripting.IDictionary = Nothing)

    LoggerInterface_log WeBeeLogger.LOG_LEVELS.levelNotice, message, context

End Function

''
' @inheritdoc
''
Private Function LoggerInterface_info(ByVal message As String, Optional ByRef context As Scripting.IDictionary = Nothing)

    LoggerInterface_log WeBeeLogger.LOG_LEVELS.levelInfo, message, context

End Function

''
' @inheritdoc
''
Private Function LoggerInterface_diagnostic(ByVal message As String, Optional ByRef context As Scripting.IDictionary = Nothing)

    LoggerInterface_log WeBeeLogger.LOG_LEVELS.levelDiagnostic, message, context

End Function
