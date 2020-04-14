VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileHandler"
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
' @var String logFilePath path to file where logs will be saved,
'                         by default "%USERPROFILE%\VBA_Logs\yyyymmdd_vba.log" where:
'                           * yyyy is 4 digit year,
'                           * mm is 2 digit month number,
'                           * dd is 2 digit day number.
''
Private logFilePath As String

''
' @var Scripting.FileSystemObject fs
''
Private fs As New Scripting.FileSystemObject

''
' @var WeBeeLogger.FormatterInterface formatter
''
Private formatter As WeBeeLogger.FormatterInterface

''
' @inheritdoc
''
Private Sub Class_Initialize()

    logFilePath = fs.BuildPath(fs.BuildPath(VBA.Environ$("USERPROFILE"), "VBA_Logs"), VBA.format$(Now(), "yyyymmdd_vba.log"))

End Sub

''
' Sets parameters to class instance.
'
' @param String filePath [opt] path to file where log records will be stored
'
' @return WeBeeLogger.HandlerInterface
''
Public Function construct( _
    Optional ByVal filePath As String = vbNullString _
) As WeBeeLogger.HandlerInterface

    If (vbNullString <> VBA.Trim$(filePath)) Then
        logFilePath = filePath
    End If

    Set construct = Me

End Function

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

    If (Not fs.FileExists(logFilePath)) Then createLogPath logFilePath

    With fs.OpenTextFile(logFilePath, ForAppending, True, TristateMixed)
        .WriteLine formatter.format(lr)
        .Close
    End With

End Function

''
' Creates missing folder structure for provided path.
'
' @param String pathToCreate
''
Private Function createLogPath(ByVal pathToCreate As String)

    If (Not fs.FolderExists(fs.GetParentFolderName(pathToCreate))) Then
        createLogPath fs.GetParentFolderName(pathToCreate)
        fs.CreateFolder fs.GetParentFolderName(pathToCreate)
    End If

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
