VERSION 5.00
Begin VB.Form frmLoadArchtype 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Load Archtype"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox lstArchtypes 
      Height          =   2985
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmLoadArchtype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdLoad_Click()
    Dim iIndex As Integer
    Dim icharnum As Integer

    For iIndex = LBound(Characters) To UBound(Characters)
        If Characters(iIndex).bInUse = False Then
            icharnum = iIndex
            Exit For
        End If
    Next
    
    Characters(icharnum) = LoadArchtypeFromFile(lstArchtypes.Text)
    Characters(icharnum).bInUse = True
  
    CreateCharacter icharnum, Characters(icharnum)
    
    Unload Me

End Sub

Private Sub cmdPreview_Click()
    ShowStats LoadArchtypeFromFile(lstArchtypes.Text)
End Sub

Private Sub Form_Load()
    Dim stRead As String

    On Error GoTo ErrorLabel
    'Debug
    Open App.Path & "\archtypes\index.dat" For Input As #1
    Do
        
        Input #1, stRead
        If stRead <> "" Then
            lstArchtypes.AddItem stRead
        End If
            
    Loop While stRead <> ""
    
ErrorLabel:

Close #1
    
End Sub
