VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Shadowrun Combat Assistant"
   ClientHeight    =   6045
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstInitiative 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   4680
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame fraImages 
      Caption         =   "Images.. this should be hidden"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   615
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   3
         Left            =   5160
         Picture         =   "frmMain.frx":0000
         Stretch         =   -1  'True
         Tag             =   "Decker"
         Top             =   120
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   37
         Left            =   4320
         Picture         =   "frmMain.frx":4942
         Stretch         =   -1  'True
         Tag             =   "Detective"
         Top             =   120
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   36
         Left            =   7680
         Picture         =   "frmMain.frx":9758
         Stretch         =   -1  'True
         Tag             =   "Dragon"
         Top             =   1560
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   35
         Left            =   6840
         Picture         =   "frmMain.frx":17342
         Stretch         =   -1  'True
         Tag             =   "Elven Decker"
         Top             =   1560
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   34
         Left            =   6000
         Picture         =   "frmMain.frx":1B5BC
         Stretch         =   -1  'True
         Tag             =   "Ganger2"
         Top             =   1560
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   33
         Left            =   5160
         Picture         =   "frmMain.frx":1FB66
         Stretch         =   -1  'True
         Tag             =   "Ganger1"
         Top             =   1560
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   32
         Left            =   4320
         Picture         =   "frmMain.frx":24308
         Stretch         =   -1  'True
         Tag             =   "Lonestar1"
         Top             =   1560
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   31
         Left            =   4320
         Picture         =   "frmMain.frx":28DFE
         Stretch         =   -1  'True
         Tag             =   "Guard1"
         Top             =   3000
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   30
         Left            =   5160
         Picture         =   "frmMain.frx":2EBC8
         Stretch         =   -1  'True
         Tag             =   "Lonestar2"
         Top             =   3000
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   29
         Left            =   6000
         Picture         =   "frmMain.frx":349A6
         Stretch         =   -1  'True
         Tag             =   "Guard2"
         Top             =   3000
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   28
         Left            =   6840
         Picture         =   "frmMain.frx":3962C
         Stretch         =   -1  'True
         Tag             =   "Corperate Mage"
         Top             =   3000
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   27
         Left            =   120
         Picture         =   "frmMain.frx":3DD46
         Stretch         =   -1  'True
         Tag             =   "Guard3"
         Top             =   3000
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   26
         Left            =   120
         Picture         =   "frmMain.frx":43AA8
         Stretch         =   -1  'True
         Tag             =   "Corperate Hitman"
         Top             =   4440
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   25
         Left            =   6840
         Picture         =   "frmMain.frx":49A76
         Stretch         =   -1  'True
         Tag             =   "Company Man"
         Top             =   4440
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   24
         Left            =   6000
         Picture         =   "frmMain.frx":4DF58
         Stretch         =   -1  'True
         Tag             =   "Street Mage F"
         Top             =   4440
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   23
         Left            =   5160
         Picture         =   "frmMain.frx":52E5A
         Stretch         =   -1  'True
         Tag             =   "Street Mage"
         Top             =   4440
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   22
         Left            =   4320
         Picture         =   "frmMain.frx":58180
         Stretch         =   -1  'True
         Tag             =   "Lonestar3"
         Top             =   4440
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   21
         Left            =   3480
         Picture         =   "frmMain.frx":5C8E2
         Stretch         =   -1  'True
         Tag             =   "Orc Merc"
         Top             =   120
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   20
         Left            =   3480
         Picture         =   "frmMain.frx":622E4
         Stretch         =   -1  'True
         Tag             =   "Guard5"
         Top             =   1560
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   19
         Left            =   3480
         Picture         =   "frmMain.frx":67AB6
         Stretch         =   -1  'True
         Tag             =   "Rigger"
         Top             =   3000
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   18
         Left            =   3480
         Picture         =   "frmMain.frx":6C804
         Stretch         =   -1  'True
         Tag             =   "Street Mage 2"
         Top             =   4440
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   17
         Left            =   2640
         Picture         =   "frmMain.frx":72286
         Stretch         =   -1  'True
         Tag             =   "Street Samuri"
         Top             =   120
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   16
         Left            =   2640
         Picture         =   "frmMain.frx":782D0
         Stretch         =   -1  'True
         Tag             =   "Street Mage 3"
         Top             =   1560
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   15
         Left            =   2640
         Picture         =   "frmMain.frx":7D01E
         Stretch         =   -1  'True
         Tag             =   "Rocker"
         Top             =   3000
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   14
         Left            =   2640
         Picture         =   "frmMain.frx":811E8
         Stretch         =   -1  'True
         Tag             =   "Shaman 3"
         Top             =   4440
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   13
         Left            =   1800
         Picture         =   "frmMain.frx":870A2
         Stretch         =   -1  'True
         Tag             =   "Shaman 2"
         Top             =   120
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   12
         Left            =   1800
         Picture         =   "frmMain.frx":8CB84
         Stretch         =   -1  'True
         Tag             =   "Mage"
         Top             =   1560
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   11
         Left            =   1800
         Picture         =   "frmMain.frx":916DE
         Stretch         =   -1  'True
         Tag             =   "Chick"
         Top             =   3000
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   10
         Left            =   1800
         Picture         =   "frmMain.frx":9587C
         Stretch         =   -1  'True
         Tag             =   "Mercenary"
         Top             =   4440
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   9
         Left            =   960
         Picture         =   "frmMain.frx":9A7D6
         Stretch         =   -1  'True
         Tag             =   "Bum"
         Top             =   120
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   8
         Left            =   960
         Picture         =   "frmMain.frx":9E7D0
         Stretch         =   -1  'True
         Tag             =   "Shaman 1"
         Top             =   1560
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   7
         Left            =   960
         Picture         =   "frmMain.frx":A3C92
         Stretch         =   -1  'True
         Tag             =   "Wannabe"
         Top             =   3000
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   6
         Left            =   960
         Picture         =   "frmMain.frx":A7E84
         Stretch         =   -1  'True
         Tag             =   "Troll Merc"
         Top             =   4440
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   5
         Left            =   120
         Picture         =   "frmMain.frx":AD53E
         Stretch         =   -1  'True
         Tag             =   "Orc Businessman"
         Top             =   120
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   4
         Left            =   120
         Picture         =   "frmMain.frx":B2A70
         Stretch         =   -1  'True
         Tag             =   "Dwarf Mercenary"
         Top             =   1560
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   1
         Left            =   6840
         Picture         =   "frmMain.frx":B631A
         Stretch         =   -1  'True
         Tag             =   "Cowboy"
         Top             =   120
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   2
         Left            =   6000
         Picture         =   "frmMain.frx":BB004
         Stretch         =   -1  'True
         Tag             =   "Cowgirl"
         Top             =   120
         Width           =   810
      End
      Begin VB.Image imgCharacter 
         Height          =   1365
         Index           =   0
         Left            =   7680
         Picture         =   "frmMain.frx":BE96E
         Stretch         =   -1  'True
         Tag             =   "Guy from a club"
         Top             =   120
         Width           =   810
      End
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label lblHealth 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Image imgActive 
      Height          =   1095
      Index           =   0
      Left            =   2400
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&Xit"
      End
   End
   Begin VB.Menu mnuNPCs 
      Caption         =   "&NPCs"
      Begin VB.Menu mnuNPCsBlank 
         Caption         =   "Add &Blank..."
      End
      Begin VB.Menu mnuNPCsArchtype 
         Caption         =   "Add &Archetype..."
      End
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "Hidden"
      Visible         =   0   'False
      Begin VB.Menu mnuHiddenName 
         Caption         =   "*NAME*"
      End
      Begin VB.Menu mnuHiddenSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHiddenView 
         Caption         =   "View Stats..."
      End
      Begin VB.Menu mnuHiddenEdit 
         Caption         =   "Edit Stats..."
      End
      Begin VB.Menu mnuHiddenTakeMissileDamage 
         Caption         =   "Take Missile Damage..."
      End
      Begin VB.Menu mnuHiddenDoMissileDamage 
         Caption         =   "Do Missile Damage..."
      End
      Begin VB.Menu mnuHiddenMelee 
         Caption         =   "Melee Combat..."
      End
      Begin VB.Menu mnuHiddenViewInventory 
         Caption         =   "View Inventory..."
      End
      Begin VB.Menu mnuHiddenRemove 
         Caption         =   "Remove"
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Begin VB.Menu mnuActionsRollini 
         Caption         =   "&Roll initiative"
      End
      Begin VB.Menu mnuActionsNewini 
         Caption         =   "&New initiative turn"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MousePositionOnImageX(1 To 1000) As Single
Dim MousePositionOnImageY(1 To 1000) As Single

Private Sub Form_DragDrop(source As Control, X As Single, Y As Single)
    imgActive(source.Index).Visible = False
    lblName(source.Index).Visible = False
    lblHealth(source.Index).Visible = False
      
    X = X - MousePositionOnImageX(source.Index)
    Y = Y - MousePositionOnImageY(source.Index)
    
    
    source.Left = X
    source.Top = Y
    
    lblHealth(source.Index).Left = X
    lblName(source.Index).Left = X
    
    lblHealth(source.Index).Top = Y - lblHealth(source.Index).Height
    lblName(source.Index).Top = Y - lblHealth(source.Index).Height - lblName(source.Index).Height

    
    imgActive(source.Index).Visible = True
    lblName(source.Index).Visible = True
    lblHealth(source.Index).Visible = True
    

End Sub

Private Sub imgActive_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then
        MousePositionOnImageX(Index) = X
        MousePositionOnImageY(Index) = Y
    ElseIf Button = 2 Then
        mnuHiddenName.Caption = Characters(Index).stName
        gCurrent = Index
        PopupMenu mnuHidden, , imgActive(Index).Left + X, imgActive(Index).Top + Y
    End If
    
End Sub

Private Sub imgActive_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePositionOnImageX(Index) = X
    MousePositionOnImageY(Index) = Y
End Sub

Private Sub lblHealth_DragDrop(Index As Integer, source As Control, X As Single, Y As Single)

    imgActive(source.Index).Visible = False
    lblName(source.Index).Visible = False
    lblHealth(source.Index).Visible = False

    X = X - MousePositionOnImageX(source.Index)
    Y = Y - MousePositionOnImageY(source.Index)


    X = X + lblHealth(Index).Left
    Y = Y + lblHealth(Index).Top
    
    source.Left = X
    source.Top = Y
    
    lblHealth(source.Index).Left = X
    lblName(source.Index).Left = X
    
    lblHealth(source.Index).Top = Y - lblHealth(source.Index).Height
    lblName(source.Index).Top = Y - lblHealth(source.Index).Height - lblName(source.Index).Height

    imgActive(source.Index).Visible = True
    lblName(source.Index).Visible = True
    lblHealth(source.Index).Visible = True


End Sub

Private Sub lblName_DragDrop(Index As Integer, source As Control, X As Single, Y As Single)
    imgActive(source.Index).Visible = False
    lblName(source.Index).Visible = False
    lblHealth(source.Index).Visible = False

    X = X - MousePositionOnImageX(source.Index)
    Y = Y - MousePositionOnImageY(source.Index)

    X = X + lblName(Index).Left
    Y = Y + lblName(Index).Top
    
    source.Left = X
    source.Top = Y
    
    lblHealth(source.Index).Left = X
    lblName(source.Index).Left = X
    
    lblHealth(source.Index).Top = Y - lblHealth(source.Index).Height
    lblName(source.Index).Top = Y - lblHealth(source.Index).Height - lblName(source.Index).Height

    imgActive(source.Index).Visible = True
    lblName(source.Index).Visible = True
    lblHealth(source.Index).Visible = True


End Sub
Private Sub imgActive_DragDrop(Index As Integer, source As Control, X As Single, Y As Single)
    imgActive(source.Index).Visible = False
    lblName(source.Index).Visible = False
    lblHealth(source.Index).Visible = False

    
    X = X - MousePositionOnImageX(source.Index)
    Y = Y - MousePositionOnImageY(source.Index)
    
    X = X + imgActive(Index).Left
    Y = Y + imgActive(Index).Top
    
    source.Left = X
    source.Top = Y
    
    lblHealth(source.Index).Left = X
    lblName(source.Index).Left = X
    
    lblHealth(source.Index).Top = Y - lblHealth(source.Index).Height
    lblName(source.Index).Top = Y - lblHealth(source.Index).Height - lblName(source.Index).Height

    imgActive(source.Index).Visible = True
    lblName(source.Index).Visible = True
    lblHealth(source.Index).Visible = True


End Sub


Private Sub mnuActionsRollini_Click()
    'This will roll the initiative for each active npc
    Dim iIndex As Integer
    Dim iIndex2 As Integer
    Dim iCharIndex As Integer
    
    Dim iInitiatives(1 To 1000) As Integer
    Dim stIniChars(1 To 1000) As String
    Dim bIniGone(1 To 1000) As Boolean
    
    
    iCharIndex = 1
    
    lstInitiative.Visible = False
    lstInitiative.Clear
    
    For iIndex = 1 To 1000
        If Characters(iIndex).bInUse = True Then
            iInitiatives(iCharIndex) = Roll(Characters(iIndex).iInitiative, 6) + Characters(iIndex).iReaction
            stIniChars(iCharIndex) = Characters(iIndex).stName
            iCharIndex = iCharIndex + 1
        End If
    Next
    

    Dim iHighIndex As Integer
    
    'Now, our list is from 1 to iCharIndex
    'I'm going to use a VERY bad sort algorithm because the max
    'it'll EVER have to sort is 1000, but in general it'll be
    'less than 10
    
    For iIndex = 1 To iCharIndex - 1
        iHighIndex = 1
        For iIndex2 = 2 To iCharIndex - 1
            If (bIniGone(iIndex2) = False) And (iInitiatives(iIndex2) >= iInitiatives(iHighIndex)) Then
                iHighIndex = iIndex2
            End If
        Next
        
        bIniGone(iHighIndex) = True
        lstInitiative.AddItem iInitiatives(iHighIndex) & " - " & stIniChars(iHighIndex), lstInitiative.ListCount
    Next
       
    lstInitiative.Visible = True
    
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuHiddenDoMissileDamage_Click()
    Load frmMissileAttack
    
    frmMissileAttack.Show vbModal
End Sub

Private Sub mnuHiddenEdit_Click()
    EditNPC gCurrent
End Sub

Private Sub mnuHiddenRemove_Click()
    Characters(gCurrent).bInUse = False
    Unload imgActive(gCurrent)
    Unload lblHealth(gCurrent)
    Unload lblName(gCurrent)
End Sub

Private Sub mnuHiddenTakeMissileDamage_Click()
    Load frmTakeMissileDamage
    
    frmTakeMissileDamage.txtArmour = Characters(gCurrent).iBArmour
    frmTakeMissileDamage.txtBody = Characters(gCurrent).iBody

    frmTakeMissileDamage.Show vbModal
End Sub

Private Sub mnuHiddenView_Click()
    ShowStats Characters(gCurrent)
End Sub

Private Sub mnuHiddenViewInventory_Click()
    MsgBox "The character's inventory is:" & vbCrLf & vbCrLf & Characters(gCurrent).stInventory
    
End Sub

Private Sub mnuNPCsArchtype_Click()
    frmLoadArchtype.Show vbModeless
End Sub

Private Sub mnuNPCsBlank_Click()
    frmBlankNPC.Show vbModeless
End Sub
