VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vb5 function enhancement"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show in Debug Window"
      Height          =   540
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   2715
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************'
'*      Created by Michael Canejo    *'
'*    Email: mikecanejo@hotmail.com  *'
'*           AIM: Mikey3dd           *'
'*************************************'
'* Name: vbpVB5.vbp   April 18, 2003 *'
'*************************************'

'Small demonstration of the functions and
'their syntax. Enjoy :)

Option Explicit

Private Sub cmdShow_Click()

    Dim arrArray As Variant
    
    Debug.Print Chr(32)
    Debug.Print "  -------------------"
    Debug.Print "->[Function Examples]<-"
    Debug.Print "  -------------------"
    
    Debug.Print Chr(32)
    Debug.Print "Round(99.65941,2) = " & Round(99.65941, 2)
    
    arrArray = Split("T e s t")
    
    Debug.Print Chr(32)
    Debug.Print "Split(" & Chr(34) & "T e s t" & Chr(34) & ") = " _
    & arrArray(0) & arrArray(1) & arrArray(2) & arrArray(3)
    
    Debug.Print Chr(32)
    Debug.Print "Join(" & Chr(34) & "arrArray, " & Chr(34) & "," _
    & Chr(34) & ") = " & Join(arrArray, ",")
    
    Debug.Print Chr(32)
    Debug.Print "InStrRev(" & Chr(34) & "Testing123" & Chr(34) & ", " _
    & Chr(34) & "123" & Chr(34) & ") = " & InStrRev("Testing123", "123")
    
    Debug.Print Chr(32)
    Debug.Print "Replace(" & Chr(34) & "Testing123" & Chr(34) & ", " _
    & Chr(34) & "ing123" & Chr(34) & ", " & Chr(34) & Chr(34) & ") = " & _
    Replace("Testing123", "ing123", "")

    Debug.Print Chr(32)
    Debug.Print "StrReverse(" & Chr(34) & "Created By Mike Canejo" _
    & Chr(34) & ") = " & StrReverse("Created By Mike Canejo")
    
    Debug.Print Chr(32)

End Sub
