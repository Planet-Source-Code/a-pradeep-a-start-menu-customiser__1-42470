VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2850
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6015
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2715
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   5985
      Begin VB.CommandButton Command1 
         Caption         =   "Candy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.Image imgLogo 
         Height          =   945
         Index           =   1
         Left            =   360
         Picture         =   "frmSplash.frx":0000
         Stretch         =   -1  'True
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright :- Freeware"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Caption         =   "Company :- NanoSoft"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         Caption         =   "Warning:- I am Fully Responsible for the damage you are about to cause !"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   2400
         Width           =   5415
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Policy 2.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4680
         TabIndex        =   4
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Win9x,Nt,2K,Xp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         TabIndex        =   5
         Top             =   1560
         Width           =   2340
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Madness...!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   1320
         TabIndex        =   7
         Top             =   840
         Width           =   3525
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2160
         TabIndex        =   6
         Top             =   240
         Width           =   1050
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
MsgBox "This is the second release and some of the features have been stripped off to make space for new options,    For any suggestions mail me at wolfeinstein13@yahoo.com, and one more thing double click on the image if you like !", vbOKOnly, "Coded By Mad Coder"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.title
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub imgLogo_Click(Index As Integer)
MsgBox "Welcome to the secret area, try moving the main form", vbOKOnly
Unload Me
End Sub

Private Sub lblCompany_Click()
  Unload Me
End Sub

Private Sub lblCompanyProduct_Click()
  Unload Me
End Sub

Private Sub lblCopyright_Click()
  Unload Me
End Sub

Private Sub lblPlatform_Click()
  Unload Me
End Sub

Private Sub lblProductName_Click()
  Unload Me
End Sub

Private Sub lblVersion_Click()
  Unload Me
End Sub

Private Sub lblWarning_Click()
  Unload Me
End Sub
