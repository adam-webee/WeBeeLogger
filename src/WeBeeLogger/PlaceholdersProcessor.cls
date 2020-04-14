VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlaceholdersProcessor"
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
' This processor allow to use placeholders in log message that will be replaced
' with value provided in the log record context.
'
' Placeholder format: {contextKeyName}
'
' @example:
'   <code>
'       Dim ctx as New Scripting.Dictionary
'
'       ctx("name") = "Adam"
'
'       log.debug "Hello {name}", ctx
'   </code>
'
'   Will format log message to this: "Hello Adam"
''
Option Explicit

Implements WeBeeLogger.ProcessorInterface

''
' @inheritdoc
''
Private Function ProcessorInterface_process(ByRef lr As WeBeeLogger.LogRecordInterface)

    If (lr.context Is Nothing) Then Exit Function

    Dim context As Scripting.Dictionary
    Dim contextKey As Variant
    Dim message As String

    With lr
        message = .message
        Set context = .context
    End With

    With New VBScript_RegExp_55.RegExp
        .Global = True
        .MultiLine = True
        .IgnoreCase = False

        For Each contextKey In context.Keys
            .Pattern = "{" & VBA.CStr(contextKey) & "}"

            message = .Replace(message, context.Item(contextKey))
        Next contextKey
    End With

    lr.message = message

End Function
