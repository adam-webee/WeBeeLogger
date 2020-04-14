VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormatterInterface"
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
' Formats log record to a format accepted by handler.
'
' @warning
'   * Output parameter must be described in implementing class.
'
' @param WeBeeLogger.LogRecordInterface lr
'
' @return Variant log record formatted to a form accepted by handler
''
Public Function format(ByRef lr As WeBeeLogger.LogRecordInterface) As Variant
End Function
