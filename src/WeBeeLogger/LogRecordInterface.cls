VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogRecordInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' @return String log message
''
Public Property Get message() As String
End Property

''
' @param String value log message
''
Public Property Let message(ByVal value As String)
End Property

''
' @return String logger name
''
Public Property Get channel() As String
End Property

''
' @return Scripting.Dictionary log context data [
'   @var String [FIELD_NAME] value can be anything that can be converted to string by @see: VBA.CStr
' ]
''
Public Property Get context() As Scripting.Dictionary
End Property

''
' @return WeBeeLogger.LOG_LEVELS
''
Public Property Get level() As WeBeeLogger.LOG_LEVELS
End Property

''
' @return String textual representation of log level (e.g.: error, info, warning)
''
Public Property Get levelName() As String
End Property

''
' @return Date date and time of invoking log record
''
Public Property Get dateTime() As Date
End Property

''
' @return Variant additional data needed for log record to be used or produced by processors, formatters or handlers
''
Public Property Get extra() As Scripting.Dictionary
End Property
