VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HandlerInterface"
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
' Checks whether the given log record will be handled by this handler.
'
' @warnings
'   * True will indicate only, that log record will be passed to @see: Self.handle - but handle will decide what to do with log record;
'   * False will indicate, that log record will not by passed to @see: Self.handle method;
'
' @param WeBeeLogger.LogRecordInterface lr
'
' @return Boolean
''
Public Function isHandling(ByRef lr As WeBeeLogger.LogRecordInterface) As Boolean
End Function

''
' Handles given log record.
'
' This method shall discard records that it does not want to handle.
'
' @warnings:
'   * if method returns true then calling further handlers will be skipped;
'
' @param WeBeeLogger.LogRecordInterface lr
'
' @return Boolean
''
Public Function handle(ByRef lr As WeBeeLogger.LogRecordInterface) As Boolean
End Function

''
' Closes the handler.
''
Public Sub discard()
End Sub

''
' @return WeBeeLogger.FormatterInterface
''
Public Property Get formatter() As WeBeeLogger.FormatterInterface
End Property

''
' @param WeBeeLogger.FormatterInterface value
''
Public Property Set formatter(ByRef value As WeBeeLogger.FormatterInterface)
End Property
