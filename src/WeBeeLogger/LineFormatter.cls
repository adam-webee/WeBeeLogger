VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LineFormatter"
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

Implements WeBeeLogger.FormatterInterface

''
' @var String LINE_LOG_PATTERN log pattern that will be used by formatter
''
Private Const LINE_LOG_PATTERN As String = "[%datetime%] %channel%.%level_name%: ""%message%"" [%context%]"

''
' @inheritdoc
''
Private Function FormatterInterface_format(ByRef lr As WeBeeLogger.LogRecordInterface) As Variant

    Dim formattedLog As String

    formattedLog = LINE_LOG_PATTERN

    With lr
        formattedLog = VBA.Replace$(formattedLog, "%datetime%", VBA.format$(.dateTime, "yyyy\/mm\/dd hh\:nn\:ss"))
        formattedLog = VBA.Replace$(formattedLog, "%channel%", .channel)
        formattedLog = VBA.Replace$(formattedLog, "%level_name%", .levelName)
        formattedLog = VBA.Replace$(formattedLog, "%message%", .message)
        If (.extra.Exists("context")) Then
            formattedLog = VBA.Replace$(formattedLog, "%context%", .extra.Item("context"))
        Else
            formattedLog = VBA.Replace$(formattedLog, "%context%", "-")
        End If
    End With

    FormatterInterface_format = formattedLog

End Function
