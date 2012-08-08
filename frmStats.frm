VERSION 5.00
Begin VB.Form frmStats 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stats for .."
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraWeapon 
      Caption         =   "Weapon/Armour"
      Height          =   1575
      Left            =   1680
      TabIndex        =   31
      Top             =   2280
      Width           =   1575
      Begin VB.TextBox txtDamage 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "4L"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtbArmour 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "0"
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtiArmour 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "0"
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "Weapon Damage"
         Height          =   495
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Armour"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "/"
         Height          =   255
         Left            =   720
         TabIndex        =   35
         Top             =   1080
         Width           =   135
      End
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   120
      Width           =   3135
   End
   Begin VB.Frame fraPrimaryStats 
      Caption         =   "Primary Stats"
      Height          =   1695
      Left            =   120
      TabIndex        =   15
      Top             =   495
      Width           =   2535
      Begin VB.TextBox txtWillpower 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtIntelligence 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtCharisma 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtStrength 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtQuickness 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtBody 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Intelligence"
         Height          =   255
         Left            =   840
         TabIndex        =   27
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Willpower"
         Height          =   255
         Left            =   1680
         TabIndex        =   26
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Charisma"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Strength"
         Height          =   255
         Left            =   1680
         TabIndex        =   24
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Quickness"
         Height          =   255
         Left            =   840
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Body"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame fraOtherStats 
      Caption         =   "Other Stats"
      Height          =   1695
      Left            =   2880
      TabIndex        =   7
      Top             =   495
      Width           =   1935
      Begin VB.TextBox txtEssence 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtReaction 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtMagic 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtInitiative 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Reaction"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Essence"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Magic"
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Ini. Dice"
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame fraImage 
      Caption         =   "Image"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   2295
      Width           =   1335
      Begin VB.Image imgPreview 
         Height          =   1215
         Left            =   240
         Stretch         =   -1  'True
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraSkills 
      Caption         =   "Skills"
      Height          =   1695
      Left            =   4920
      TabIndex        =   3
      Top             =   495
      Width           =   2055
      Begin VB.TextBox txtSkills 
         BackColor       =   &H80000004&
         Height          =   1095
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "[skillname] [value][enter]"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame fraInventory 
      Caption         =   "Inventory"
      Height          =   1575
      Left            =   3480
      TabIndex        =   1
      Top             =   2295
      Width           =   2055
      Begin VB.TextBox txtInventory 
         BackColor       =   &H80000004&
         Height          =   1215
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   135
      Width           =   615
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub
