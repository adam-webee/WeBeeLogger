VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConstructableInterface"
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
' Class constructor for Logger instances - as VBA has no constructor that can take parameters.
'
' @warnings
'   * This method must be called as the first method after object instantiation with New
'
' @param String name [opt] name of the logger
'
' @return void
''
Public Function construct(Optional ByVal name As String = vbNullString) As WeBeeLogger.LoggerInterface
End Function

''
' Register handlers to be used by Logger instance.
'
' @param WeBeeLogger.HandlerInterface value
'
' @return void
''
Public Sub registerHandler(ByRef value As WeBeeLogger.HandlerInterface)
End Sub

''
' Register log record processors to be used by Logger instance.
'
' By design it is allowed to register as many processors as needed.
'
' @warning:
'   * Registering order will reflect processors execution order.
'
' @param WeBeeLogger.ProcessorInterface value
'
' @return void
''
Public Sub registerProcessor(ByRef value As WeBeeLogger.ProcessorInterface)
End Sub
