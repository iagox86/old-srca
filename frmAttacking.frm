VERSION 5.00
Begin VB.Form frmMissileAttack 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Missile Attack"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   1695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   735
   End
   Begin VB.Frame fraVerify 
      Caption         =   "Please Verify..."
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1455
      Begin VB.TextBox txtDamageLevel 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   600
         TabIndex        =   3
         Top             =   1080
         Width           =   255
      End
      Begin VB.TextBox txtDamage 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtFirearms 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Damage"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Firearms Skill"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox txtTN 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "4"
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Note: Changing the skill/damage values here will not change them permenantly."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Target Number"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmMissileAttack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'To do damage, we basically roll his firearms skill against
    'the target number, then display the successes and power of
    'the gun to the user
    
    Dim iSuccesses As Integer
    iSuccesses = RollDice(txtFirearms.Text, txtTN.Text)
    
    MsgBox "He scored " & iSuccesses & "." & vbCrLf & vbCrLf & _
    "The power of his weapon is " & txtDamage.Text & txtDamageLevel.Text & "." _
    , , "Results"
    
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim iIndex As Integer
    Dim stSkills() As String
    Dim stTemp() As String
    
    On Error Resume Next
    
    txtTN.SelStart = 0
    txtTN.SelLength = Len(txtTN.Text)

    txtDamage.Text = Characters(gCurrent).iDamage
    txtDamageLevel.Text = Characters(gCurrent).iDamageLevel
    
    If Characters(gCurrent).stSkills <> "" Then
        stSkills = Split(Characters(gCurrent).stSkills, vbCrLf)
        For iIndex = 0 To UBound(stSkills)
            stTemp = Split(stSkills(iIndex), " ")
            If UCase(stTemp(0)) = "FIREARMS" Or UCase(stTemp(0)) = "FIREARM" Then
                txtFirearms.Text = stTemp(1)
                Exit Sub
            End If
        Next
    End If
    
    MsgBox "Error! No firearm skill.  Skill set to 0.", vbInformation, "Error!"
    txtFirearms.Text = "0"

End Sub

Private Sub txtTN_GotFocus()
    SelAll txtTN
End Sub


Private Sub txtDamage_GotFocus()
    SelAll txtDamage
End Sub


Private Sub txtDamageLevel_GotFocus()
    SelAll txtDamageLevel
End Sub


Private Sub txtFirearms_GotFocus()
    SelAll txtFirearms
End Sub

