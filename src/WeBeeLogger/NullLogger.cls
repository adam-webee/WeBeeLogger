VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NullLogger"
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
'
'
' This is NullLogger that implements WeBeeLogger.LoggerInterface that can be used whenever
' logger object is needed but logging is not necessary.
'
' All methods are dummy and do absolutly nothing.
''
Option Explicit

Implements WeBeeLogger.ConstructableInterface
Implements WeBeeLogger.LoggerInterface

''
' @inheritdoc
''
Private Function ConstructableInterface_construct(Optional ByVal name As String = vbNullString) As LoggerInterface

    ConstructableInterface_construct = Me

End Function

''
' @inheritdoc
''
Private Sub ConstructableInterface_registerHandler(value As HandlerInterface)
End Sub

''
' @inheritdoc
''
Private Sub ConstructableInterface_registerProcessor(value As ProcessorInterface)
End Sub

''
' @inheritdoc
''
Private Function LoggerInterface_log(ByVal level As WeBeeLogger.LOG_LEVELS, ByVal message As String, Optional ByRef context As Scripting.IDictionary = Nothing)
End Function

''
' @inheritdoc
''
Private Function LoggerInterface_emergency(ByVal message As String, Optional ByRef context As Scripting.IDictionary = Nothing)
End Function

''
' @inheritdoc
''
Private Function LoggerInterface_alert(ByVal message As String, Optional ByRef context As Scripting.IDictionary = Nothing)
End Function

''
' @inheritdoc
''
Private Function LoggerInterface_critical(ByVal message As String, Optional ByRef context As Scripting.IDictionary = Nothing)
End Function

''
' @inheritdoc
''
Private Function LoggerInterface_error(ByVal message As String, Optional ByRef context As Scripting.IDictionary = Nothing)
End Function

''
' @inheritdoc
''
Private Function LoggerInterface_warning(ByVal message As String, Optional ByRef context As Scripting.IDictionary = Nothing)
End Function

''
' @inheritdoc
''
Private Function LoggerInterface_notice(ByVal message As String, Optional ByRef context As Scripting.IDictionary = Nothing)
End Function

''
' @inheritdoc
''
Private Function LoggerInterface_info(ByVal message As String, Optional ByRef context As Scripting.IDictionary = Nothing)
End Function

''
' @inheritdoc
''
Private Function LoggerInterface_diagnostic(ByVal message As String, Optional ByRef context As Scripting.IDictionary = Nothing)
End Function
