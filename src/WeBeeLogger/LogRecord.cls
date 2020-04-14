VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogRecord"
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

Implements LogRecordInterface

''
' @var Scripting.Dictionary properties [
'   @var Date                   dateTime   log record creation date and time
'   @var String                 message    message to be logged
'   @var String                 channel    logger name, "null" by default
'   @var WeBeeLogger.LOG_LEVELS level      level of the log
'   @var Scripting.Dictionary   context    log context data [
'       @var String [KEY_NAME] value must be string or something that can by casted to string with @see: VBA.CStr
'   ]
'   @var Scripting.Dictionary   extra      additional log data to be used or produced by processors or formatters
'   @var Scripting.Dictionary   levelNames [
'       @var String [LOG_LEVEL] textual representation of log level
'   ]
' ]
''
Private properties As New Scripting.Dictionary

''
' @inheritdoc
''
Private Sub Class_Initialize()

    With properties
        .Item("dateTime") = VBA.Now()
        .Item("message") = vbNullString
        .Item("channel") = "NULL"
        .Item("level") = WeBeeLogger.LOG_LEVELS.levelDiagnostic
        Set .Item("context") = New Scripting.Dictionary
        Set .Item("extra") = New Scripting.Dictionary
        Set .Item("levelNames") = New Scripting.Dictionary
        With .Item("levelNames")
            .Item(WeBeeLogger.LOG_LEVELS.levelAlert) = "alert"
            .Item(WeBeeLogger.LOG_LEVELS.levelDiagnostic) = "debug"
            .Item(WeBeeLogger.LOG_LEVELS.levelEmergency) = "emergency"
            .Item(WeBeeLogger.LOG_LEVELS.levelError) = "error"
            .Item(WeBeeLogger.LOG_LEVELS.levelInfo) = "info"
            .Item(WeBeeLogger.LOG_LEVELS.levelNotice) = "notice"
            .Item(WeBeeLogger.LOG_LEVELS.levelWarning) = "warning"
            .Item(WeBeeLogger.LOG_LEVELS.levleCritical) = "critical"
        End With
    End With

End Sub

''
' @inheritdoc
''
Public Function construct( _
    ByVal message As String, _
    ByVal level As WeBeeLogger.LOG_LEVELS, _
    Optional channel As String = vbNullString, _
    Optional context As Scripting.IDictionary = Nothing _
) As WeBeeLogger.LogRecordInterface

    With properties
        .Item("message") = message
        .Item("level") = level
        If (vbNullString <> channel) Then .Item("channel") = channel
        If (Not context Is Nothing) Then Set .Item("context") = context
    End With

    Set construct = Me

End Function

''
' @inheritdoc
''
Private Property Get LogRecordInterface_context() As Scripting.IDictionary

    Set LogRecordInterface_context = properties.Item("context")

End Property

''
' @inheritdoc
''
Private Property Get LogRecordInterface_dateTime() As Date

    LogRecordInterface_dateTime = properties.Item("dateTime")

End Property

''
' @inheritdoc
''
Private Property Get LogRecordInterface_extra() As Scripting.IDictionary

    Set LogRecordInterface_extra = properties.Item("extra")

End Property

''
' @inheritdoc
''
Private Property Get LogRecordInterface_level() As WeBeeLogger.LOG_LEVELS

    LogRecordInterface_level = properties.Item("level")

End Property

''
' @inheritdoc
''
Private Property Get LogRecordInterface_levelName() As String

    LogRecordInterface_levelName = properties.Item("levelNames")(properties.Item("level"))

End Property

''
' @inheritdoc
''
Private Property Get LogRecordInterface_message() As String

    LogRecordInterface_message = properties.Item("message")

End Property

''
' @inheritdoc
''
Private Property Get LogRecordInterface_channel() As String

    LogRecordInterface_channel = properties.Item("channel")

End Property

''
' @inheritdoc
''
Private Property Let LogRecordInterface_message(ByVal value As String)

    properties.Item("message") = value

End Property
