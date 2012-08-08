VERSION 5.00
Begin VB.Form frmTakeMissileDamage 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Taking damage (...)"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   2730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   2040
      Width           =   735
   End
   Begin VB.Frame fraOther 
      Caption         =   "Please Verify"
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   1695
      Begin VB.TextBox txtArmour 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Text            =   "0"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtBody 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "0"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Armour"
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Body"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Info about attack"
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2415
      Begin VB.CheckBox chkAPDS 
         Caption         =   "Using APDS?"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtSuccesses 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtDamageLevel 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   480
         TabIndex        =   1
         Text            =   "L"
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox txtDamage 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Text            =   "4"
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Player Successes"
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Damage"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Note: Changing body/armour here will NOT affect them permenantly."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   2535
   End
End
Attribute VB_Name = "frmTakeMissileDamage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'When taking missile damage, this is what happens:
    '1. Find the target number to resist by subtracting ballistic armour from
    '   the gun's power (unless APDS rounds are being used, in which case it's
    '   half armour
    '2. Roll body against the target number
    '3. Subtract your successes from theirs and stage damage accordingly
    
    Dim iTN As Integer
    Dim iSuccesses As Integer
    Dim iNetSuccesses As Integer
    Dim iDamageLevel As Integer
    
    '1:
    If chkAPDS.Value Then
        iTN = Val(txtDamage.Text) - (Val(txtArmour.Text) / 2)
    Else
        iTN = Val(txtDamage.Text) - Val(txtArmour.Text)
    End If
    If iTN < 2 Then iTN = 2
    
    '2:
    iSuccesses = RollDice(Val(txtBody.Text), iTN)
    
    '3:
    iNetSuccesses = Val(txtSuccesses.Text) - iSuccesses
    iNetSuccesses = iNetSuccesses / 2
    
    txtDamageLevel.Text = UCase(txtDamageLevel.Text)
    
    'Convert the damage code to something usable
    With txtDamageLevel
        If .Text = "L" Then
            iDamageLevel = 1
        ElseIf .Text = "M" Then
            iDamageLevel = 2
        ElseIf .Text = "S" Then
            iDamageLevel = 3
        ElseIf .Text = "D" Then
            iDamageLevel = 4
        Else
            MsgBox "Bad damage level"
            Unload Me
        End If
    End With
    
    'Get the new damagelevel and make sure it's not outside the range
    iDamageLevel = iDamageLevel + iNetSuccesses
    If iDamageLevel < 0 Then
        iDamageLevel = 0
    ElseIf iDamageLevel > 4 Then
        iDamageLevel = 4
    End If
    

    If iDamageLevel = 1 Then
        Characters(gCurrent).iHealth = Characters(gCurrent).iHealth + 1
    ElseIf iDamageLevel = 2 Then
        Characters(gCurrent).iHealth = Characters(gCurrent).iHealth + 3
    ElseIf iDamageLevel = 3 Then
        Characters(gCurrent).iHealth = Characters(gCurrent).iHealth + 6
    ElseIf iDamageLevel = 4 Then
        Characters(gCurrent).iHealth = Characters(gCurrent).iHealth + 10
    End If
    
    If Characters(gCurrent).iHealth > 10 Then
        Characters(gCurrent).iHealth = 10
    End If
        
    frmMain.lblHealth(gCurrent).Caption = Characters(gCurrent).iHealth
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    txtDamage.SelStart = 0
    txtDamage.SelLength = Len(txtDamage.Text)
End Sub

Private Sub txtDamage_GotFocus()
    SelAll txtDamage
End Sub
Private Sub txtDamageLevel_GotFocus()
    SelAll txtDamageLevel
End Sub
Private Sub txtsuccesses_GotFocus()
    SelAll txtSuccesses
End Sub
Private Sub txtbody_GotFocus()
    SelAll txtBody
End Sub
Private Sub txtarmour_GotFocus()
    SelAll txtArmour
End Sub

