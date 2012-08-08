VERSION 5.00
Begin VB.Form frmBlankNPC 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add New NPC"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtEditGuy 
      Height          =   285
      Left            =   4080
      TabIndex        =   37
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame fraInventory 
      Caption         =   "Inventory"
      Height          =   1575
      Left            =   4920
      TabIndex        =   34
      Top             =   2280
      Width           =   2055
      Begin VB.TextBox txtInventory 
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame fraSkills 
      Caption         =   "Skills"
      Height          =   1695
      Left            =   4920
      TabIndex        =   32
      Top             =   480
      Width           =   2055
      Begin VB.TextBox txtSkills 
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   35
         Text            =   "frmBlankNPC.frx":0000
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "[skillname] [value][enter]"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   4920
      TabIndex        =   15
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6000
      TabIndex        =   16
      Top             =   3960
      Width           =   975
   End
   Begin VB.Frame fraWeapon 
      Caption         =   "Weapon/Armour"
      Height          =   1575
      Left            =   3240
      TabIndex        =   30
      Top             =   2280
      Width           =   1575
      Begin VB.TextBox txtiArmour 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   40
         Text            =   "0"
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtbArmour 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   39
         Text            =   "0"
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtDamage 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   13
         Text            =   "4L"
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "/"
         Height          =   255
         Left            =   720
         TabIndex        =   41
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label13 
         Caption         =   "Armour"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Weapon Damage"
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraImage 
      Caption         =   "Image"
      Height          =   1575
      Left            =   120
      TabIndex        =   29
      Top             =   2280
      Width           =   3015
      Begin VB.ListBox lstPicture 
         Height          =   1230
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
      Begin VB.Image imgPreview 
         Height          =   1215
         Left            =   2040
         Stretch         =   -1  'True
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraOtherStats 
      Caption         =   "Other Stats"
      Height          =   1695
      Left            =   2880
      TabIndex        =   24
      Top             =   480
      Width           =   1935
      Begin VB.TextBox txtInitiative 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Text            =   "1"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtMagic 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Reaction"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtReaction 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtEssence 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Text            =   "6"
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Ini. Dice"
         Height          =   255
         Left            =   1080
         TabIndex        =   28
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Magic"
         Height          =   255
         Left            =   1080
         TabIndex        =   27
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Essence"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Frame fraPrimaryStats 
      Caption         =   "Primary Stats"
      Height          =   1695
      Left            =   120
      TabIndex        =   17
      Top             =   480
      Width           =   2535
      Begin VB.TextBox txtBody 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Text            =   "0"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtQuickness 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtStrength 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtCharisma 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Text            =   "0"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtIntelligence 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtWillpower 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Body"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Quickness"
         Height          =   255
         Left            =   840
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Strength"
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Charisma"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Willpower"
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Intelligence"
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "New NPC"
      Top             =   105
      Width           =   3135
   End
   Begin VB.Label Label9 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmBlankNPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mEditNum As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim iIndex As Integer
    Dim icharnum As Integer
    
    If txtEditGuy.Text = "" Then
        For iIndex = 1 To 1000
            If Characters(iIndex).bInUse = False Then
                icharnum = iIndex
                Exit For
            End If
        Next
    Else
        icharnum = Val(txtEditGuy.Text)
        frmMain.lblName(icharnum) = txtName.Text
        
    End If
    
    Characters(icharnum) = MakeCharacterVariable( _
                                     txtName.Text, _
                                     txtBody.Text, _
                                     txtQuickness.Text, _
                                     txtStrength.Text, _
                                     txtCharisma.Text, _
                                     txtIntelligence.Text, _
                                     txtWillpower.Text, _
                                     txtEssence.Text, _
                                     txtMagic.Text, _
                                     txtReaction.Text, _
                                     txtInitiative.Text, _
                                     txtDamage.Text, _
                                     txtbArmour.Text, _
                                     txtiArmour.Text, _
                                     txtInventory.Text, _
                                     txtSkills.Text, _
                                     lstPicture.ItemData(lstPicture.ListIndex))
                    
    If txtEditGuy.Text = "" Then
        Characters(icharnum).bInUse = True
        CreateCharacter icharnum, Characters(icharnum)
    Else
        frmMain.imgActive(icharnum).Picture = frmMain.imgCharacter(lstPicture.ItemData(lstPicture.ListIndex)).Picture
    End If
    
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim temp As Character
    
    temp = MakeCharacterVariable( _
                                     txtName.Text, _
                                     txtBody.Text, _
                                     txtQuickness.Text, _
                                     txtStrength.Text, _
                                     txtCharisma.Text, _
                                     txtIntelligence.Text, _
                                     txtWillpower.Text, _
                                     txtEssence.Text, _
                                     txtMagic.Text, _
                                     txtReaction.Text, _
                                     txtInitiative.Text, _
                                     txtDamage.Text, _
                                     txtbArmour.Text, _
                                     txtiArmour.Text, _
                                     txtInventory.Text, _
                                     txtSkills.Text, _
                                     lstPicture.ItemData(lstPicture.ListIndex))
                                     
    SaveArchtypeToFile temp
    MsgBox txtName.Text & ".dat has been saved!", vbInformation, "Success!"
    
End Sub

Private Sub Form_Load()
    Dim iIndex As Integer
    Dim i As Integer

    For iIndex = frmMain.imgCharacter.LBound To frmMain.imgCharacter.UBound
        lstPicture.AddItem frmMain.imgCharacter(iIndex).Tag
        For i = 0 To lstPicture.ListCount
            If lstPicture.List(i) = frmMain.imgCharacter(iIndex).Tag Then
                lstPicture.ItemData(i) = iIndex
            End If
        Next
    Next
    
    lstPicture.ListIndex = 0
    

End Sub

Private Sub lstPicture_Click()
    imgPreview.Picture = frmMain.imgCharacter(lstPicture.ItemData(lstPicture.ListIndex)).Picture
        
End Sub

Private Sub txtName_GotFocus()
    Sel txtName
End Sub

Private Sub txtbody_GotFocus()
    Sel txtBody
End Sub

Private Sub txtQuickness_GotFocus()
    Sel txtQuickness
End Sub

Private Sub txtStrength_GotFocus()
    Sel txtStrength
End Sub

Private Sub txtCharisma_GotFocus()
    Sel txtCharisma
End Sub

Private Sub txtintelligence_GotFocus()
    Sel txtIntelligence
End Sub

Private Sub txtWillpower_GotFocus()
    Sel txtWillpower
End Sub
Private Sub txtEssence_GotFocus()
    Sel txtEssence
End Sub
Private Sub txtMagic_GotFocus()
    Sel txtMagic
End Sub
Private Sub txtReaction_GotFocus()
    Sel txtReaction
End Sub
Private Sub txtInitiative_GotFocus()
    Sel txtInitiative
End Sub

Private Sub txtDamage_GotFocus()
    Sel txtDamage
End Sub

Private Sub Sel(source As Object)
    source.SelStart = 0
    source.SelLength = Len(source.Text)
End Sub
