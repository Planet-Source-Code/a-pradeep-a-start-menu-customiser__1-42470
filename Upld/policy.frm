VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Madness for Windows  V 2.0.0"
   ClientHeight    =   3945
   ClientLeft      =   2265
   ClientTop       =   2190
   ClientWidth     =   8145
   Icon            =   "policy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   8145
   Begin MSComDlg.CommonDialog cmdlg1 
      Left            =   7200
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Apply Now"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   76
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CheckBox Check55 
      Caption         =   "Add To Context menu"
      Height          =   255
      Left            =   5760
      TabIndex        =   74
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton done 
      Caption         =   "I Want My Mommy"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   5741
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Start Menu"
      TabPicture(0)   =   "policy.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Desktop"
      TabPicture(1)   =   "policy.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Windows Explorer"
      TabPicture(2)   =   "policy.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame4 
         Caption         =   "Simple"
         Height          =   2655
         Left            =   -74640
         TabIndex        =   15
         Top             =   480
         Width           =   7335
         Begin VB.CheckBox Check59 
            Caption         =   "Select Active html ?"
            Height          =   255
            Left            =   2760
            TabIndex        =   80
            Top             =   1440
            Width           =   1935
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Bitmap for Windows Tool Bar"
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   360
            Width           =   2415
         End
         Begin VB.CheckBox Check58 
            Caption         =   "No Active Desktop"
            Height          =   255
            Left            =   2760
            TabIndex        =   78
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CheckBox Check57 
            Caption         =   "Delete VB Recent files"
            Height          =   255
            Left            =   2760
            TabIndex        =   77
            Top             =   720
            Width           =   1935
         End
         Begin VB.CheckBox Check56 
            Caption         =   "Add Log On Message"
            Height          =   255
            Left            =   2760
            TabIndex        =   75
            Top             =   360
            Width           =   2055
         End
         Begin VB.CheckBox Check54 
            Caption         =   "Clear Media Player menu"
            Height          =   375
            Left            =   120
            TabIndex        =   69
            Top             =   2160
            Width           =   2055
         End
         Begin VB.CheckBox Check53 
            Caption         =   "Clear Find Menu"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   1800
            Width           =   1455
         End
         Begin VB.CheckBox Check52 
            Caption         =   "Clear Run Menu"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Du IT !"
            Height          =   375
            Left            =   6360
            TabIndex        =   65
            Top             =   1320
            Width           =   615
         End
         Begin VB.CheckBox Check24 
            Caption         =   "Hide Drives in Explorer"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CheckBox Check23 
            Caption         =   "Small Icons For Tool Bar"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   720
            Width           =   2175
         End
         Begin VB.Frame Frame3 
            Caption         =   "Drives To Hide"
            Height          =   1815
            Left            =   2520
            TabIndex        =   37
            Top             =   240
            Width           =   4575
            Begin VB.CheckBox Check50 
               Caption         =   "Z"
               Height          =   255
               Left            =   3960
               TabIndex        =   63
               Top             =   720
               Width           =   375
            End
            Begin VB.CheckBox Check49 
               Caption         =   "Y"
               Height          =   255
               Left            =   3960
               TabIndex        =   62
               Top             =   360
               Width           =   375
            End
            Begin VB.CheckBox Check48 
               Caption         =   "X"
               Height          =   255
               Left            =   3360
               TabIndex        =   61
               Top             =   1440
               Width           =   375
            End
            Begin VB.CheckBox Check47 
               Caption         =   "W"
               Height          =   255
               Left            =   3360
               TabIndex        =   60
               Top             =   1080
               Width           =   495
            End
            Begin VB.CheckBox Check46 
               Caption         =   "V"
               Height          =   255
               Left            =   3360
               TabIndex        =   59
               Top             =   720
               Width           =   495
            End
            Begin VB.CheckBox Check45 
               Caption         =   "U"
               Height          =   255
               Left            =   3360
               TabIndex        =   58
               Top             =   360
               Width           =   375
            End
            Begin VB.CheckBox Check44 
               Caption         =   "T"
               Height          =   255
               Left            =   2760
               TabIndex        =   57
               Top             =   1440
               Width           =   375
            End
            Begin VB.CheckBox Check43 
               Caption         =   "S"
               Height          =   255
               Left            =   2760
               TabIndex        =   56
               Top             =   1080
               Width           =   375
            End
            Begin VB.CheckBox Check42 
               Caption         =   "R"
               Height          =   255
               Left            =   2760
               TabIndex        =   55
               Top             =   720
               Width           =   495
            End
            Begin VB.CheckBox Check41 
               Caption         =   "Q"
               Height          =   255
               Left            =   2760
               TabIndex        =   54
               Top             =   360
               Width           =   495
            End
            Begin VB.CheckBox Check40 
               Caption         =   "P"
               Height          =   255
               Left            =   2040
               TabIndex        =   53
               Top             =   1440
               Width           =   495
            End
            Begin VB.CheckBox Check39 
               Caption         =   "O"
               Height          =   255
               Left            =   2040
               TabIndex        =   52
               Top             =   1080
               Width           =   495
            End
            Begin VB.CheckBox Check38 
               Caption         =   "N"
               Height          =   255
               Left            =   2040
               TabIndex        =   51
               Top             =   720
               Width           =   495
            End
            Begin VB.CheckBox Check37 
               Caption         =   "M"
               Height          =   255
               Left            =   2040
               TabIndex        =   50
               Top             =   360
               Width           =   495
            End
            Begin VB.CheckBox Check36 
               Caption         =   "L"
               Height          =   255
               Left            =   1440
               TabIndex        =   49
               Top             =   1440
               Width           =   375
            End
            Begin VB.CheckBox Check35 
               Caption         =   "K"
               Height          =   255
               Left            =   1440
               TabIndex        =   48
               Top             =   1080
               Width           =   495
            End
            Begin VB.CheckBox Check34 
               Caption         =   "J"
               Height          =   255
               Left            =   1440
               TabIndex        =   47
               Top             =   720
               Width           =   495
            End
            Begin VB.CheckBox Check33 
               Caption         =   "I"
               Height          =   255
               Left            =   1440
               TabIndex        =   46
               Top             =   360
               Width           =   375
            End
            Begin VB.CheckBox Check32 
               Caption         =   "H"
               Height          =   255
               Left            =   840
               TabIndex        =   45
               Top             =   1440
               Width           =   495
            End
            Begin VB.CheckBox Check31 
               Caption         =   "G"
               Height          =   255
               Left            =   840
               TabIndex        =   44
               Top             =   1080
               Width           =   375
            End
            Begin VB.CheckBox Check30 
               Caption         =   "F"
               Height          =   255
               Left            =   840
               TabIndex        =   43
               Top             =   720
               Width           =   495
            End
            Begin VB.CheckBox Check29 
               Caption         =   "E"
               Height          =   255
               Left            =   840
               TabIndex        =   42
               Top             =   360
               Width           =   495
            End
            Begin VB.CheckBox Check28 
               Caption         =   "D"
               Height          =   255
               Left            =   240
               TabIndex        =   41
               Top             =   1440
               Width           =   495
            End
            Begin VB.CheckBox Check27 
               Caption         =   "C"
               Height          =   255
               Left            =   240
               TabIndex        =   40
               Top             =   1080
               Width           =   615
            End
            Begin VB.CheckBox Check26 
               Caption         =   "B"
               Height          =   255
               Left            =   240
               TabIndex        =   39
               Top             =   720
               Width           =   615
            End
            Begin VB.CheckBox Check25 
               Caption         =   "A"
               Height          =   255
               Left            =   240
               TabIndex        =   38
               Top             =   360
               Width           =   495
            End
         End
         Begin VB.Image Image1 
            Height          =   1680
            Left            =   4920
            Stretch         =   -1  'True
            Top             =   480
            Width           =   2160
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Rename"
         Height          =   2415
         Left            =   -74640
         TabIndex        =   14
         Top             =   600
         Width           =   7095
         Begin VB.CommandButton Command5 
            Caption         =   "Du"
            Height          =   255
            Left            =   2520
            TabIndex        =   73
            Top             =   2040
            Width           =   615
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   120
            TabIndex        =   72
            Top             =   1920
            Width           =   2175
         End
         Begin VB.CheckBox Check21 
            Caption         =   "Text in Media Player title"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   1560
            Width           =   2055
         End
         Begin VB.CheckBox Check20 
            Caption         =   "No AutoRun CD's"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   1200
            Width           =   1935
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Recycle Bin"
            Height          =   255
            Left            =   2640
            TabIndex        =   66
            Top             =   1800
            Width           =   1455
         End
         Begin VB.CheckBox Check51 
            Caption         =   "Hide Control panel  printers"
            Height          =   375
            Left            =   120
            TabIndex        =   64
            Top             =   720
            Width           =   2295
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Chicken !"
            Height          =   375
            Left            =   5880
            TabIndex        =   34
            Top             =   1680
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Du It !"
            Height          =   375
            Left            =   4800
            TabIndex        =   33
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   4800
            TabIndex        =   32
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   4800
            TabIndex        =   31
            Top             =   480
            Width           =   2055
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Schedule Tasks"
            Height          =   255
            Left            =   2640
            TabIndex        =   28
            Top             =   1440
            Width           =   1575
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Printers"
            Height          =   255
            Left            =   2640
            TabIndex        =   27
            Top             =   1080
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Dial Up Networking"
            Height          =   255
            Left            =   2640
            TabIndex        =   26
            Top             =   720
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Control Panel"
            Height          =   255
            Left            =   2640
            TabIndex        =   25
            Top             =   360
            Width           =   1335
         End
         Begin VB.CheckBox Check12 
            Caption         =   "No Desktop Icons"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Tool Tip"
            Height          =   255
            Left            =   5280
            TabIndex        =   30
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Name"
            Height          =   255
            Left            =   5280
            TabIndex        =   29
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Disable in Start Menu"
         Height          =   2535
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   7335
         Begin VB.CheckBox Check18 
            Caption         =   "Disable Printer Adding"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   2160
            Width           =   1935
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Disable Printer Deleting"
            Height          =   255
            Left            =   2760
            TabIndex        =   23
            Top             =   2160
            Width           =   2055
         End
         Begin VB.CheckBox Check19 
            Caption         =   "Disable Settings"
            Height          =   255
            Left            =   5280
            TabIndex        =   22
            Top             =   2160
            Width           =   1575
         End
         Begin VB.CheckBox Check17 
            Caption         =   "Task Bar Properties"
            Height          =   255
            Left            =   5280
            TabIndex        =   21
            Top             =   1800
            Width           =   1815
         End
         Begin VB.CheckBox Check16 
            Caption         =   "StartMenu Changing"
            Height          =   255
            Left            =   5280
            TabIndex        =   20
            Top             =   1440
            Width           =   1815
         End
         Begin VB.CheckBox Check15 
            Caption         =   "Hide My Pictures"
            Height          =   255
            Left            =   5280
            TabIndex        =   19
            Top             =   1080
            Width           =   1815
         End
         Begin VB.CheckBox Check14 
            Caption         =   "Hide My Documents"
            Height          =   255
            Left            =   5280
            TabIndex        =   18
            Top             =   720
            Width           =   1935
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Multiple Coloumns"
            Height          =   255
            Left            =   5280
            TabIndex        =   17
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Disable Shut Down"
            Height          =   255
            Left            =   2760
            TabIndex        =   13
            Top             =   1800
            Width           =   1695
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Disable Log Off"
            Height          =   255
            Left            =   2760
            TabIndex        =   12
            Top             =   1440
            Width           =   1815
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Disable Run Menu"
            Height          =   255
            Left            =   2760
            TabIndex        =   11
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Disable Folder Options"
            Height          =   255
            Left            =   2760
            TabIndex        =   10
            Top             =   720
            Width           =   2175
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Recent Documents History"
            Height          =   255
            Left            =   2760
            TabIndex        =   9
            Top             =   360
            Width           =   2295
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Hide Start Menu SubFolders"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   1800
            Width           =   2295
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Hide Windows Update"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1440
            Width           =   1935
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Hide Find"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Hide Favorites"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Hide Documents"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As Long
Dim title, mesg As String
Private Sub Check1_Click()
Madness "NoRecentDocsMenu", Check1
End Sub
Public Function Madness(name As String, obj As Object)
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", name, obj.Value
'SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", name, obj.Value
End Function
Private Sub Check10_Click()
Madness "NoClose", Check10
End Sub
Private Sub Check11_Click()
If Check11.Value = 1 Then
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "StartMenuScrollPrograms", "YES"
Else
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "StartMenuScrollPrograms", "NO"
End If
End Sub
Private Sub Check12_Click()
Madness "NoDesktop", Check12
End Sub
Private Sub Check13_Click()
Madness "NoDeletePrinter", Check13
End Sub
Private Sub Check14_Click()
Madness "NoSMMyDocs", Check14
End Sub
Private Sub Check15_Click()
Madness "NoSMMyPictures", Check15
End Sub
Private Sub Check16_Click()
Madness "NoChangeStartMenu", Check16
End Sub
Private Sub Check17_Click()
Madness "NoSetTaskbar", Check17
End Sub
Private Sub Check18_Click()
Madness "NoAddPrinter", Check18
End Sub
Private Sub Check19_Click()
MsgBox "Under R&D"
End Sub
Private Sub Check2_Click()
Madness "NoFavoritesMenu", Check2
End Sub

Private Sub Check20_Click()
Madness "NoDriveTypeAutoRun", Check20
End Sub

Private Sub Check21_Click()
CreateKey "HKEY_CURRENT_USER\Software\Policies\Microsoft\WindowsMediaPlayer"
t = GetStringValue("HKEY_CURRENT_USER\Software\Policies\Microsoft\WindowsMediaPlayer", "Titlebar")
If t <> "Error" Then
Text3.Text = t
End If
If Check21.Value = 0 Then
Text3.Visible = False
Command5.Visible = False
End If
If Check21.Value = 1 Then
Text3.Visible = True
Command5.Visible = True
End If
End Sub


Private Sub Check23_Click()
If Check23.Value = 1 Then
temp = SetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\SmallIcons", "SmallIcons", "Yes")
ElseIf Check23.Value = 0 Then
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\SmallIcons", "SmallIcons", "No"
End If
End Sub
Private Sub Check24_Click()
'this is a lil complicated cos the number of drives are crutial for this module
If Check24.Value = 1 Then
Frame3.Visible = True
Command4.Visible = True
Image1.Visible = False
Check56.Visible = False
Check57.Visible = False
Check58.Visible = False
Check59.Visible = False
Else
Frame3.Visible = False
Check56.Visible = True
Check57.Visible = True
Check58.Visible = True
Check59.Visible = True
Image1.Visible = True
Command4.Visible = False
End If
End Sub
Private Sub Check25_Click()
If Check25.Value = 1 Then
temp = temp + 1000000
End If
If Check25.Value = 0 Then
temp = 1000000 - temp
End If
End Sub

Private Sub Check26_Click()
If Check26.Value = 1 Then
temp = temp + 200000
End If
If Check26.Value = 0 Then
temp = temp - 200000
End If
End Sub
Private Sub Check27_Click()
If Check27.Value = 1 Then
temp = temp + 400000
End If
If Check27.Value = 0 Then
temp = temp - 400000
End If
End Sub

Private Sub Check28_Click()
If Check28.Value = 1 Then
temp = temp + 800000
End If
If Check28.Value = 0 Then
temp = temp - 800000
End If
End Sub

Private Sub Check29_Click()
If Check29.Value = 1 Then
temp = temp + 10000000
End If
If Check29.Value = 0 Then
temp = temp - 10000000
End If
End Sub

Private Sub Check30_Click()
If Check30.Value = 1 Then
temp = temp + 20000000
End If
If Check30.Value = 0 Then
temp = temp - 20000000
End If
End Sub

Private Sub Check3_Click()
Madness "NoFind", Check3
End Sub

Private Sub Check31_Click()
If Check31.Value = 1 Then
temp = temp + 40000000
End If
If Check31.Value = 0 Then
temp = temp - 40000000
End If
End Sub

Private Sub Check32_Click()
If Check32.Value = 1 Then
temp = temp + 80000000
End If
If Check32.Value = 0 Then
temp = temp - 80000000
End If
End Sub

Private Sub Check33_Click()
If Check33.Value = 1 Then
temp = temp + 10000
End If
If Check33.Value = 0 Then
temp = temp - 10000
End If
End Sub

Private Sub Check34_Click()
If Check34.Value = 1 Then
temp = temp + 20000
End If
If Check34.Value = 0 Then
temp = temp - 20000
End If
End Sub

Private Sub Check35_Click()
If Check35.Value = 1 Then
temp = temp + 40000
End If
If Check35.Value = 0 Then
temp = temp - 40000
End If
End Sub

Private Sub Check36_Click()
If Check36.Value = 1 Then
temp = temp + 80000
End If
If Check36.Value = 0 Then
temp = temp - 80000
End If
End Sub

Private Sub Check37_Click()
If Check37.Value = 1 Then
temp = temp + 100000
End If
If Check37.Value = 0 Then
temp = temp - 100000
End If
End Sub

Private Sub Check38_Click()
If Check38.Value = 1 Then
temp = temp + 200000
End If
If Check38.Value = 0 Then
temp = temp - 200000
End If
End Sub

Private Sub Check39_Click()
If Check39.Value = 1 Then
temp = temp + 400000
End If
If Check39.Value = 0 Then
temp = temp - 400000
End If
End Sub

Private Sub Check4_Click()
Madness "NoWindowsUpdate", Check4
End Sub

Private Sub Check40_Click()
If Check40.Value = 1 Then
temp = temp + 800000
End If
If Check40.Value = 0 Then
temp = temp - 800000
End If
End Sub

Private Sub Check41_Click()
If Check41.Value = 1 Then
temp = temp + 100
End If
If Check41.Value = 0 Then
temp = temp - 100
End If
End Sub

Private Sub Check42_Click()
If Check42.Value = 1 Then
temp = temp + 200
End If
If Check42.Value = 0 Then
temp = temp - 200
End If
End Sub

Private Sub Check43_Click()
If Check43.Value = 1 Then
temp = temp + 400
End If
If Check43.Value = 0 Then
temp = temp - 400
End If
End Sub

Private Sub Check44_Click()
If Check44.Value = 1 Then
temp = temp + 800
End If
If Check43.Value = 0 Then
temp = temp - 800
End If
End Sub

Private Sub Check45_Click()
If Check45.Value = 1 Then
temp = temp + 1000
End If
If Check45.Value = 0 Then
temp = temp - 1000
End If
End Sub

Private Sub Check46_Click()
If Check46.Value = 1 Then
temp = temp + 2000
End If
If Check46.Value = 0 Then
temp = temp - 2000
End If
End Sub

Private Sub Check47_Click()
If Check47.Value = 1 Then
temp = temp + 4000
End If
If Check47.Value = 0 Then
temp = temp - 4000
End If
End Sub

Private Sub Check48_Click()
If Check48.Value = 1 Then
temp = temp + 8000
End If
If Check49.Value = 0 Then
temp = temp + 8000
End If
End Sub

Private Sub Check49_Click()
If Check49.Value = 1 Then
temp = temp + 10
End If
If Check49.Value = 0 Then
temp = temp - 10
End If
End Sub

Private Sub Check5_Click()
Madness "NoStartMenuSubFolders", Check5
End Sub

Private Sub Check50_Click()
If Check50.Value = 1 Then
temp = temp + 20
End If
If Check50.Value = 0 Then
temp = temp - 20
End If
End Sub

Private Sub Check51_Click()
Madness "NoSetFolders", Check51
End Sub

Private Sub Check52_Click()
RegDeleteKey &H80000001, "Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU"
On Error Resume Next
End Sub

Private Sub Check53_Click()
RegDeleteKey &H80000001, "Software\Microsoft\Windows\CurrentVersion\Explorer\Doc Find Spec MRU"
On Error Resume Next
End Sub

Private Sub Check54_Click()
RegDeleteKey &H80000001, "Software\Microsoft\MediaPlayer\Player\RecentURLList"
On Error Resume Next
End Sub

Private Sub Check55_Click()
If Check55.Value = 1 Then
CreateKey ("HKEY_CLASSES_ROOT\Folder\shell\Policy")
CreateKey ("HKEY_CLASSES_ROOT\Folder\shell\Policy\command")
SetStringValue "HKEY_CLASSES_ROOT\Folder\shell\Policy\command", "", App.Path & "\" & App.EXEName & ".exe"
End If
If Check55.Value = 0 Then
RegDeleteKey &H80000000, "Folder\shell\Policy\command"
RegDeleteKey &H80000000, "Folder\shell\Policy"
End If
On Error Resume Next
End Sub
Private Sub Check56_Click()

If Check56.Value = 1 Then
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Winlogon", "LegalNoticeCaption", InputBox("Enter text")
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Winlogon", "LegalNoticeText", InputBox("Enter Title")
End If
If Check56.Value = 0 Then
RegDeleteKey &H80000002, "Software\Microsoft\Windows\CurrentVersion\Winlogon"
On Error Resume Next
End If
End Sub

Private Sub Check57_Click()
If Check57.Value = 1 Then
RegDeleteKey &H80000001, "Software\Microsoft\Visual Basic\6.0\RecentFiles"
End If
End Sub

Private Sub Check58_Click()
If Check58.Value = 1 Then
Madness "NoActiveDesktop", Check58
End If
End Sub

Private Sub Check59_Click()
If Check59.Value = 1 Then
cmdlg1.FileName = "*.html"
cmdlg1.ShowOpen
End If
If cmdlg1.FileName <> "*.html" Then
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Desktop\SafeMode\General", "Wallpaper", cmdlg1.FileName
End If
End Sub
Private Sub Check6_Click()
Madness "NoRecentDocsHistory", Check6
End Sub
Private Sub Check7_Click()
Madness "NoFolderOptions", Check7
End Sub
Private Sub Check8_Click()
Madness "NoRun", Check8
End Sub
Private Sub Check9_Click()
Madness "NoLogOff", Check9
End Sub
Private Sub Command1_Click()
If Option1.Value = True Then
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{21EC2020-3AEA-1069-A2DD-08002B30309D}", "", Text1.Text
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{21EC2020-3AEA-1069-A2DD-08002B30309D}", "InfoTip", Text2.Text
End If
If Option2.Value = True Then
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{992cffa0-f557-101a-88ec-00dd010ccc48}", "", Text1.Text
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{992cffa0-f557-101a-88ec-00dd010ccc48}", "InfoTip", Text2.Text
End If
If Option3.Value = True Then
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{2227A280-3AEA-1069-A2DE-08002B30309D}", "", Text1.Text
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{2227A280-3AEA-1069-A2DE-08002B30309D}", "InfoTip", Text2.Text
End If
If Option4.Value = True Then
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{D6277990-4C6A-11CF-8D87-00AA0060F5BF}", "", Text1.Text
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{D6277990-4C6A-11CF-8D87-00AA0060F5BF}", "InfoTip", Text2.Text
End If
If Option5.Value = True Then
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "", Text1.Text
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "InfoTip", Text2.Text
End If
Command1.Visible = False
Command3.Visible = False
Text1.Visible = False
Text2.Visible = False
Label1.Visible = False
Label2.Visible = False
End Sub
Private Sub Command2_Click()
frmSplash.Show
End Sub
Private Sub Command3_Click()
Command1.Visible = False
Command3.Visible = False
Text1.Visible = False
Text2.Visible = False
Label1.Visible = False
Label2.Visible = False
End Sub
Private Sub Command4_Click()
'SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDrives", Str(temp)
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDrives", Str(temp)
Command4.Visible = False
Frame3.Visible = False
Image1.Visible = True
Check56.Visible = True
Check57.Visible = True
Check58.Visible = True
Check59.Visible = True
End Sub
Private Sub Command5_Click()
SetStringValue "HKEY_CURRENT_USER\Software\Policies\Microsoft\WindowsMediaPlayer", "Titlebar", Text3.Text
Text3.Visible = False
Command5.Visible = False
End Sub
Private Sub Command6_Click()
If vbOK = MsgBox("Some require Windows to restart hit okay for restart?", vbOKCancel, "Madness") Then
ExitWindowsEx 3, 1
End If
End Sub
Private Sub done_Click()
'End
Unload Me
End Sub
Private Sub Form_Activate()
If App.PrevInstance Then
MsgBox "Copy Already Running !", vbExclamation
End
End If
CreateKey "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
CreateKey ("HKEY_CLASSES_ROOT\Folder\shell\Policy")
CreateKey ("HKEY_CLASSES_ROOT\Folder\shell\Policy\command")
SetStringValue "HKEY_CLASSES_ROOT\Folder\shell\Policy\command", "", App.Path & "\" & App.EXEName & ".exe"
'to check user

res = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoActiveDesktop")
If res = "Error" Or res = "0" Then
Check58.Value = 0
ElseIf res = "1" Then
Check58.Value = 1
End If

res = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\SmallIcons", "")
If res = "Yes" Then
Check23.Value = 1
ElseIf res = "No" Then
Check23.Value = 0
End If

res = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDesktop")
If res = "Error" Or res = "0" Then
Check12.Value = 0
End If
If res = "1" Then
Check12.Value = 1
End If

res = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDrivetypeautorun")
If res = "Error" Or res = "0" Then
Check20.Value = 0
End If
If res = "1" Then
Check20.Value = 1
End If

res = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "Nosetfolders")
If res <> "Error" Or res = "0" Then
Check51.Value = 0
End If
If res = "1" Then
Check51.Value = 1
End If
End Sub
Private Sub Form_Load()
Check55.Value = 1
Text3.Visible = False
Command5.Visible = False
Command4.Visible = False
temp = 0
Frame3.Visible = False
Text1.Visible = False
Text2.Visible = False
Label1.Visible = False
Label2.Visible = False
Command1.Visible = False
Command3.Visible = False
res = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "StartMenuScrollPrograms")
If res <> "Error" Then
    If res = "YES" Then
    Check11.Value = 1
    Else
    Check11.Value = 0
    End If
End If

res = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu")
If res <> "Error" Or res = "0" Then
Check1.Value = 0
End If
If res = "1" Then
Check1.Value = 1
End If

res = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu")
If res <> "Error" Or res = "0" Then
Check2.Value = 0
End If
If res = "1" Then
Check2.Value = 1
End If


res = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind")
If res <> "Error" Or res = "0" Then
Check3.Value = 0
End If
If res = "1" Then
Check3.Value = 1
End If

res = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWindowsUpdate")
If res <> "Error" Or res = "0" Then
Check4.Value = 0
End If
If res = "1" Then
Check4.Value = 1
End If

res = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoStartMenuSubFolders")
If res <> "Error" Or res = "0" Then
Check5.Value = 0
End If
If res = "1" Then
Check5.Value = 1
End If

res = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoAddPrinter")
If res <> "Error" Or res = "0" Then
Check18.Value = 0
End If
If res = "1" Then
Check18.Value = 1
End If

res = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsHistory")
If res <> "Error" Or res = "0" Then
Check6.Value = 0
End If
If res = "1" Then
Check6.Value = 1
End If


res = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions")
If res <> "Error" Or res = "0" Then
Check7.Value = 0
End If
If res = "1" Then
Check7.Value = 1
End If

res = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun")
If res <> "Error" Or res = "0" Then
Check8.Value = 0
End If
If res = "1" Then
Check8.Value = 1
End If


res = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogOff")
If res <> "Error" Or res = "0" Then
Check9.Value = 0
End If
If res = "1" Then
Check9.Value = 1
End If

res = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose")
If res <> "Error" Or res = "0" Then
Check10.Value = 0
End If
If res = "1" Then
Check10.Value = 1
End If

res = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDeletePrinter")
If res <> "Error" Or res = "0" Then
Check13.Value = 0
End If
If res = "1" Then
Check13.Value = 1
End If

res = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSMMyDocs")
If res <> "Error" Or res = "0" Then
Check14.Value = 0
End If
If res = "1" Then
Check14.Value = 1
End If

res = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSMMyPictures")
If res <> "Error" Or res = "0" Then
Check15.Value = 0
End If
If res = "1" Then
Check15.Value = 1
End If

res = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoChangeStartMenu")
If res <> "Error" Or res = "0" Then
Check16.Value = 0
End If
If res = "1" Then
Check16.Value = 1
End If

res = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetTaskbar")
If res <> "Error" Or res = "0" Then
Check17.Value = 0
End If
If res = "1" Then
Check17.Value = 1
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
temp = Form1.Top
For i = Form1.Height To 1 Step -2
Form1.Height = i
Form1.Top = temp + 1
temp = temp + 1
Next i
j = Form1.Left
For i = Form1.Width To 1 Step -2
Form1.Width = i
Form1.Left = j
j = j + 1
Next i
End
End Sub

Private Sub Image1_DblClick()
t = MsgBox("Do you want to save this image to Disk ? ", vbOKCancel)
If t = vbOK Then
SavePicture Image1.Picture, "c:\candy3.bmp"
MsgBox "Saved as c:\candy3.bmp"
End If
End Sub
Private Sub Option1_Click()
If Option1.Value = True Then
Text1.Text = GetStringValue("HKEY_CLASSES_ROOT\CLSID\{21EC2020-3AEA-1069-A2DD-08002B30309D}", "")
Text2.Text = GetStringValue("HKEY_CLASSES_ROOT\CLSID\{21EC2020-3AEA-1069-A2DD-08002B30309D}", "InfoTip")
Text1.Visible = True
Text2.Visible = True
Command1.Visible = True
Command3.Visible = True
Label1.Visible = True
Label2.Visible = True
End If
End Sub
Private Sub Option2_Click()
If Option2.Value = True Then
Text1.Text = GetStringValue("HKEY_CLASSES_ROOT\CLSID\{992cffa0-f557-101a-88ec-00dd010ccc48}", "")
Text2.Text = GetStringValue("HKEY_CLASSES_ROOT\CLSID\{992cffa0-f557-101a-88ec-00dd010ccc48}", "InfoTip")
Text1.Visible = True
Text2.Visible = True
Command1.Visible = True
Command3.Visible = True
Label1.Visible = True
Label2.Visible = True
End If
End Sub
Private Sub Option3_Click()
If Option3.Value = True Then
Text1.Text = GetStringValue("HKEY_CLASSES_ROOT\CLSID\{2227A280-3AEA-1069-A2DE-08002B30309D}", "")
Text2.Text = GetStringValue("HKEY_CLASSES_ROOT\CLSID\{2227A280-3AEA-1069-A2DE-08002B30309D}", "InfoTip")
Text1.Visible = True
Text2.Visible = True
Command1.Visible = True
Command3.Visible = True
Label1.Visible = True
Label2.Visible = True
End If
End Sub
Private Sub Option4_Click()
If Option4.Value = True Then
Text1.Text = GetStringValue("HKEY_CLASSES_ROOT\CLSID\{D6277990-4C6A-11CF-8D87-00AA0060F5BF}", "")
Text2.Text = GetStringValue("HKEY_CLASSES_ROOT\CLSID\{D6277990-4C6A-11CF-8D87-00AA0060F5BF}", "InfoTip")
Text1.Visible = True
Text2.Visible = True
Command1.Visible = True
Command3.Visible = True
Label1.Visible = True
Label2.Visible = True
End If
End Sub
Private Sub Option5_Click()
If Option5.Value = True Then
Text1.Text = GetStringValue("HKEY_CLASSES_ROOT\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "")
Text2.Text = GetStringValue("HKEY_CLASSES_ROOT\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "InfoTip")
Text1.Visible = True
Text2.Visible = True
Command1.Visible = True
Command3.Visible = True
Label1.Visible = True
Label2.Visible = True
End If
End Sub

Private Sub Option6_Click()
If Option6.Value = True Then
cmdlg1.FileName = "*.bmp"
cmdlg1.ShowOpen
If cmdlg1.FileName <> "*.bmp" Then
SetStringValue "HKEY_USERS\.DEFAULT\Software\Microsoft\Internet Explorer\Toolbar", "BackBitmapShell", cmdlg1.FileName
End If
If cmdlg1.FileName = "*.bmp" Then
Option6.Value = False
End If
End If
End Sub

