VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "MS Visual C Runtime QSort with random values"
      Height          =   465
      Left            =   375
      TabIndex        =   1
      Top             =   1875
      Width           =   3765
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   375
      TabIndex        =   0
      Top             =   150
      Width           =   3765
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim hFile       As Long
    Dim pFnc        As Long
    Dim hCb         As Long
    Dim lngArr()    As Long
    Dim i           As Long
    
    ' 256 random values
    ReDim lngArr(255) As Long
    
    Randomize
    
    For i = 0 To UBound(lngArr)
        lngArr(i) = Rnd() * 1000
    Next
    
    ' get the function pointer of qsort in msvcrt.dll
    pFnc = GetProcAddressEx("msvcrt.dll", "qsort")
    
    ' create a cdecl callback wich points to QSortCallback
    ' the callback has 2 parameters
    hCb = CreateCdeclCbWrap(AddressOf QSortCallback, 2)
    ' call qsort
    CallCdecl pFnc, VarPtr(lngArr(0)), UBound(lngArr) + 1, 4, hCb
    ' destroy the cdecl callback to release the allocated memory
    DestroyDeclCbWrap hCb
    
    ' show the sorted array
    List1.Clear
    For i = 0 To UBound(lngArr)
        List1.AddItem lngArr(i)
    Next
End Sub
