VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImmediateHandler"
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

Implements WeBeeLogger.HandlerInterface

''
' @var WeBeeLogger.FormatterInterface formatter
''
Private formatter As WeBeeLogger.FormatterInterface

''
' @inheritdoc
''
Private Property Get HandlerInterface_formatter() As WeBeeLogger.FormatterInterface

    Set HandlerInterface_formatter = formatter

End Property

''
' @inheritdoc
''
Private Property Set HandlerInterface_formatter(ByRef value As WeBeeLogger.FormatterInterface)

    Set formatter = value

End Property

''
' @inheritdoc
''
Private Function HandlerInterface_handle(ByRef lr As WeBeeLogger.LogRecordInterface) As Boolean

    Debug.Print VBA.CStr(formatter.format(lr))

End Function

''
' @inheritdoc
''
Private Function HandlerInterface_isHandling(ByRef lr As WeBeeLogger.LogRecordInterface) As Boolean

    HandlerInterface_isHandling = True

End Function

''
' @inheritdoc
''
Private Sub HandlerInterface_discard()
End Sub
