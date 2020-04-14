VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoggerInterface"
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
'
'
' This interface is highly inspired and compliant as much as possible to
' one defined for PHP in PSR-3: @see: https://www.php-fig.org/psr/psr-3/
'
' @warnings
'   * Context variable must be one level deep where all keys are strings and all values are strings
'     or convertible to strings by @see: VBA.CStr;
''
Option Explicit

Public Enum LOG_LEVELS
    levelEmergency = 0
    levelAlert = 1
    levleCritical = 2
    levelError = 3
    levelWarning = 4
    levelNotice = 5
    levelInfo = 6
    levelDiagnostic = 7
End Enum

''
' Logs with an arbitrary level.
'
' @param WeBeeLogger.LOG_LEVELS level
' @param String               message
' @param Scripting.Dictionary context [opt]
'
' @return void
''
Public Function log(ByVal level As WeBeeLogger.LOG_LEVELS, ByVal message As String, Optional ByRef context As Scripting.Dictionary = Nothing)
End Function

''
' System is unusable.
'
' @param String               message
' @param Scripting.Dictionary context [opt]
'
' @return void
''
Public Function emergency(ByVal message As String, Optional ByRef context As Scripting.Dictionary = Nothing)
End Function

''
' Action must be taken immediately.
'
' Example: Entire solution does not work, database is unavailable, etc.
'
' @param String               message
' @param Scripting.Dictionary context [opt]
'
' @return void
''
Public Function alert(ByVal message As String, Optional ByRef context As Scripting.Dictionary = Nothing)
End Function

''
' Critical conditions.
'
' Example: Application component unavailable, unexpected exception.
'
' @param String               message
' @param Scripting.Dictionary context [opt]
'
' @return void
''
Public Function critical(ByVal message As String, Optional ByRef context As Scripting.Dictionary = Nothing)
End Function

''
' Runtime errors that do not require immediate action but should typically be logged and monitored.
'
' @param String               message
' @param Scripting.Dictionary context [opt]
'
' @return void
''
Public Function error(ByVal message As String, Optional ByRef context As Scripting.Dictionary = Nothing)
End Function

''
' Exceptional occurrences that are not errors.
'
' Example: Use of depreciated API, poor use of an API, undesirable things that are not necessarily wrong.
'
' @param String               message
' @param Scripting.Dictionary context [opt]
'
' @return void
''
Public Function warning(ByVal message As String, Optional ByRef context As Scripting.Dictionary = Nothing)
End Function

''
' Normal but significant events.
'
' @param String               message
' @param Scripting.Dictionary context [opt]
'
' @return void
''
Public Function notice(ByVal message As String, Optional ByRef context As Scripting.Dictionary = Nothing)
End Function

''
' Interesting events.
'
' @param String               message
' @param Scripting.Dictionary context [opt]
'
' @return void
''
Public Function info(ByVal message As String, Optional ByRef context As Scripting.Dictionary = Nothing)
End Function

''
' Detailed debug information.
'
' @warning
'   As in VBA keyword debug is reserved word diagnostic is introduced instead.
'
' @param String               message
' @param Scripting.Dictionary context [opt]
'
' @return void
''
Public Function diagnostic(ByVal message As String, Optional ByRef context As Scripting.Dictionary = Nothing)
End Function
