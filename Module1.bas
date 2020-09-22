Attribute VB_Name = "Module1"
Option Explicit

' for some help on qsort google for it
Public Function QSortCallback(ByRef value1 As Long, ByRef value2 As Long) As Long
    QSortCallback = value1 - value2
End Function
