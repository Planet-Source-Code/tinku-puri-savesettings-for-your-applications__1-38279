VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings(Without API Call and Without INI File.Very Easy.)"
   ClientHeight    =   6135
   ClientLeft      =   2580
   ClientTop       =   1335
   ClientWidth     =   6615
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6615
   Begin VB.Frame MainFrame 
      Height          =   5775
      Left            =   45
      TabIndex        =   0
      Top             =   300
      Width           =   6525
      Begin VB.Frame Frame1 
         Caption         =   "Check Boxes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   135
         TabIndex        =   18
         Top             =   2715
         Width           =   1485
         Begin VB.CheckBox Check3 
            Caption         =   "Check3"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   105
            TabIndex        =   21
            Top             =   990
            Width           =   1260
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Check2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   105
            TabIndex        =   20
            Top             =   630
            Width           =   1290
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   105
            TabIndex        =   19
            Top             =   300
            Width           =   1305
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Option Button"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   1650
         TabIndex        =   14
         Top             =   2715
         Width           =   1590
         Begin VB.OptionButton Option3 
            Caption         =   "Option3"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   17
            Top             =   945
            Width           =   1365
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Option2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   16
            Top             =   615
            Width           =   1440
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   105
            TabIndex        =   15
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Click here to change Form's caption"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   900
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2085
         Width           =   4725
      End
      Begin VB.Frame Frame3 
         Caption         =   "Back Color"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   915
         TabIndex        =   8
         Top             =   645
         Width           =   2295
         Begin VB.HScrollBar HScroll1 
            Height          =   195
            LargeChange     =   10
            Left            =   120
            Max             =   255
            TabIndex        =   11
            Top             =   360
            Width           =   2055
         End
         Begin VB.HScrollBar HScroll2 
            Height          =   195
            LargeChange     =   10
            Left            =   120
            Max             =   255
            TabIndex        =   10
            Top             =   615
            Width           =   2055
         End
         Begin VB.HScrollBar HScroll3 
            Height          =   195
            LargeChange     =   10
            Left            =   120
            Max             =   255
            TabIndex        =   9
            Top             =   885
            Width           =   2055
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   1605
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Text            =   "Form1.frx":0442
         Top             =   4635
         Width           =   3270
      End
      Begin VB.Frame Frame4 
         Caption         =   "Fonts"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   3345
         TabIndex        =   5
         Top             =   2730
         Width           =   3015
         Begin VB.ListBox List1 
            Height          =   1035
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   240
            Width           =   2760
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Fore Color"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   3330
         TabIndex        =   1
         Top             =   645
         Width           =   2295
         Begin VB.HScrollBar HScroll4 
            Height          =   195
            LargeChange     =   10
            Left            =   120
            Max             =   255
            TabIndex        =   4
            Top             =   885
            Width           =   2055
         End
         Begin VB.HScrollBar HScroll5 
            Height          =   195
            LargeChange     =   10
            Left            =   120
            Max             =   255
            TabIndex        =   3
            Top             =   615
            Width           =   2055
         End
         Begin VB.HScrollBar HScroll6 
            Height          =   195
            LargeChange     =   10
            Left            =   120
            Max             =   255
            TabIndex        =   2
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Save Settings DEMO... By TINKU"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1365
         TabIndex        =   22
         Top             =   180
         Width           =   3480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Type Your Text Here."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1620
         TabIndex        =   13
         Top             =   4335
         Width           =   1830
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "email:Hollowman_Tinku@yahoo.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1290
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   30
      Width           =   4005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kl
Dim kl2
'this settings demo show how to save your customize settings withour api without ini
'file you have just use VB's inbuilt function Savesettings and Getsettings with this you
'can do easyly save your settings
'=======================================
 'Intruction How use this functions
'=======================================
    '-------------For Save Settings---------
        'Put This function Form_Unload event or Command_Click Event Where ever you want, you can put this function
        '1).Parameter Of This Functions
            'SaveSetting (AppName as String,Section as String,Key as String,Setting as string)
            '-> AppName as string
                'in this parameter you have to spacify Application name('if you want save your programm's settings then you have to put APP.EXENAME)
            '->Section as string
                'in this parameter you have to specify you section in which section you have to save Settings."
                'For Example you have to save your Settings to Section "SETTING".you can give any Section Name
            '-> Key as strings
                'in this parameter you have to specify you Control's Settings Key
                'As Example You Can give any name to key
            '->Setting as string
               'In this you have to specify your settings
               'As Example you have to save Form's BackColor then write Form1.backcolor
               'in this way you can save any control's value to the settings
    '---------------For Get Settings-----------
        'put this function in Form_Load Event or Command_Click event Where ever you want,you  can put this function
        'GetSetting(AppName as String,Section as String,Key as strings ,[Default] as String)
            '->AppName as String
                'in this parameter you have to spacify Application name('if you want Get your programm's settings then you have to put APP.EXENAME)
            '->Section as String
                'in this parameter you have to specify From Where you have to get your settings
                'as example if You had save you settings under "SETTING" Section Then Write This Section Name
            '->Key as String
                'Please Specify your key (Which you had in savesettings)
                'As Example you had saveyour form'top in "FORMTOP" then Write "FORMTOP"
'=====================================================
    'FOR MORE HELP SEE WHOLE CODE.
    'MAIL ME FOR GET MORE CODES LIKE.....ABOUT MULTIMEDIA,API CALLS,SKIN REFRENCE
    'E -MAIL: HOLLOWMAN_TINKU@ YAHOO.COM
'=====================================================
Private Sub Command1_Click()
    klp = InputBox("Enter String For Form Caption")
    Me.Caption = klp
    Command1.Caption = klp
End Sub

Private Sub Form_Initialize()
    If App.PrevInstance = True Then
        MsgBox "App Already Running.", vbOKOnly + vbCritical + vbApplicationModal
        End
    End If
End Sub

Private Sub Form_Load()
    For X = 0 To Screen.FontCount - 1
        List1.AddItem Screen.Fonts(X)
    Next
    Me.Top = GetSetting(App.EXEName, "Settings", "METOP", (Screen.Height - Me.Height) / 2)
    Me.Left = GetSetting(App.EXEName, "Settings", "MELEFT", (Screen.Width - Me.Width) / 2)
    Me.Caption = GetSetting(App.EXEName, "Settings", "MECAPTION", "Settings(Without API Call and Without INI File.Very Easy.)")
    Check1.Value = GetSetting(App.EXEName, "Settings", "Check1", 0)
    Check2.Value = GetSetting(App.EXEName, "Settings", "Check2", 1)
    Check3.Value = GetSetting(App.EXEName, "Settings", "Check3", 0)
    Option1.Value = GetSetting(App.EXEName, "Settings", "Option1", 0)
    Option2.Value = GetSetting(App.EXEName, "Settings", "Option2", 0)
    Option3.Value = GetSetting(App.EXEName, "Settings", "Option3", 0)
    HScroll1.Value = GetSetting(App.EXEName, "Settings", "Hscroll1", 113)
    HScroll2.Value = GetSetting(App.EXEName, "Settings", "Hscroll2", 120)
    HScroll3.Value = GetSetting(App.EXEName, "Settings", "Hscroll3", 134)
    HScroll4.Value = GetSetting(App.EXEName, "Settings", "Hscroll4", 255)
    HScroll5.Value = GetSetting(App.EXEName, "Settings", "Hscroll5", 255)
    HScroll6.Value = GetSetting(App.EXEName, "Settings", "Hscroll6", 255)
    Text1.Text = GetSetting(App.EXEName, "Settings", "TEXT", "Settings Demo By Tinku Without API Calls and Without INI File.")
    On Error GoTo tinku
    List1.ListIndex = GetSetting(App.EXEName, "Settings", "LIST", List1.List(List1.Text = "Tahoma"))
tinku:
    Exit Sub
End Sub

Private Sub Form_Resize()
'    MainFrame.Top = (Me.Height - MainFrame.Height) / 2
'    MainFrame.Left = (Me.Width - MainFrame.Width) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    SaveSetting App.EXEName, "Settings", "Check1", Check1.Value
    SaveSetting App.EXEName, "Settings", "Check2", Check2.Value
    SaveSetting App.EXEName, "Settings", "Check3", Check3.Value
    SaveSetting App.EXEName, "Settings", "Option1", Option1.Value
    SaveSetting App.EXEName, "Settings", "Option2", Option2.Value
    SaveSetting App.EXEName, "Settings", "Option3", Option3.Value
    SaveSetting App.EXEName, "Settings", "Hscroll1", HScroll1.Value
    SaveSetting App.EXEName, "Settings", "Hscroll2", HScroll2.Value
    SaveSetting App.EXEName, "Settings", "Hscroll3", HScroll3.Value
    SaveSetting App.EXEName, "Settings", "Hscroll4", HScroll4.Value
    SaveSetting App.EXEName, "Settings", "Hscroll5", HScroll5.Value
    SaveSetting App.EXEName, "Settings", "Hscroll6", HScroll6.Value
    SaveSetting App.EXEName, "Settings", "METOP", Me.Top
    SaveSetting App.EXEName, "Settings", "MELEFT", Me.Left
    SaveSetting App.EXEName, "Settings", "MECAPTION", Me.Caption
    SaveSetting App.EXEName, "Settings", "TEXT", Text1.Text
    SaveSetting App.EXEName, "Settings", "LIST", List1.ListIndex
End Sub

Private Sub HScroll1_Change()
    kl = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
    Me.BackColor = kl
    Command1.BackColor = kl
    Frame1.BackColor = kl
    Frame2.BackColor = kl
    Frame3.BackColor = kl
    Frame4.BackColor = kl
    Frame5.BackColor = kl
    Option1.BackColor = kl
    Option2.BackColor = kl
    Option3.BackColor = kl
    Check1.BackColor = kl
    Check2.BackColor = kl
    Check3.BackColor = kl
    Label1.BackColor = kl
    Text1.BackColor = kl
    List1.BackColor = kl
    MainFrame.BackColor = kl
End Sub

Private Sub HScroll1_Scroll()
    kl = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
    Me.BackColor = kl
    Label1.BackColor = kl
    Command1.BackColor = kl
    Frame1.BackColor = kl
    Frame2.BackColor = kl
    Frame3.BackColor = kl
    Frame4.BackColor = kl
    Frame5.BackColor = kl
    Option1.BackColor = kl
    Option2.BackColor = kl
    Option3.BackColor = kl
    Check1.BackColor = kl
    Check2.BackColor = kl
    Check3.BackColor = kl
    Text1.BackColor = kl
    List1.BackColor = kl
    MainFrame.BackColor = kl
End Sub
Private Sub HScroll2_Change()
    kl = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
    Me.BackColor = kl
    Label1.BackColor = kl
    Command1.BackColor = kl
    Frame1.BackColor = kl
    Frame2.BackColor = kl
    Frame3.BackColor = kl
    Frame4.BackColor = kl
    Frame5.BackColor = kl
    Option1.BackColor = kl
    Option2.BackColor = kl
    Option3.BackColor = kl
    Check1.BackColor = kl
    Check2.BackColor = kl
    Check3.BackColor = kl
    Text1.BackColor = kl
    List1.BackColor = kl
    MainFrame.BackColor = kl
End Sub

Private Sub HScroll2_Scroll()
    kl = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
    Me.BackColor = kl
    Label1.BackColor = kl
    Command1.BackColor = kl
    Frame1.BackColor = kl
    Frame2.BackColor = kl
    Frame3.BackColor = kl
    Frame4.BackColor = kl
    Frame5.BackColor = kl
    Option1.BackColor = kl
    Option2.BackColor = kl
    Option3.BackColor = kl
    Check1.BackColor = kl
    Check2.BackColor = kl
    Check3.BackColor = kl
    Text1.BackColor = kl
    List1.BackColor = kl
    MainFrame.BackColor = kl
End Sub
Private Sub HScroll3_Change()
    kl = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
    Me.BackColor = kl
    Label1.BackColor = kl
    Command1.BackColor = kl
    Frame1.BackColor = kl
    Frame2.BackColor = kl
    Frame3.BackColor = kl
    Frame4.BackColor = kl
    Frame5.BackColor = kl
    Option1.BackColor = kl
    Option2.BackColor = kl
    Option3.BackColor = kl
    Check1.BackColor = kl
    Check2.BackColor = kl
    Check3.BackColor = kl
    Text1.BackColor = kl
    List1.BackColor = kl
    MainFrame.BackColor = kl
End Sub

Private Sub HScroll3_Scroll()
    kl = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
    Me.BackColor = kl
    Label1.BackColor = kl
    Command1.BackColor = kl
    Frame1.BackColor = kl
    Frame2.BackColor = kl
    Frame3.BackColor = kl
    Frame4.BackColor = kl
    Frame5.BackColor = kl
    Option1.BackColor = kl
    Option2.BackColor = kl
    Option3.BackColor = kl
    Check1.BackColor = kl
    Check2.BackColor = kl
    Check3.BackColor = kl
    Text1.BackColor = kl
    List1.BackColor = kl
    MainFrame.BackColor = kl
End Sub

Private Sub HScroll6_Change()
    kl2 = RGB(HScroll6.Value, HScroll5.Value, HScroll4.Value)
    Check1.ForeColor = kl2
    Check2.ForeColor = kl2
    Check3.ForeColor = kl2
    Option1.ForeColor = kl2
    Option2.ForeColor = kl2
    Option3.ForeColor = kl2
    List1.ForeColor = kl2
    Text1.ForeColor = kl2
    Frame1.ForeColor = kl2
    Frame2.ForeColor = kl2
    Frame3.ForeColor = kl2
    Frame4.ForeColor = kl2
    Frame5.ForeColor = kl2
    Label1.ForeColor = kl2
    Label2.ForeColor = kl2
End Sub
Private Sub HScroll5_Change()
    kl2 = RGB(HScroll6.Value, HScroll5.Value, HScroll4.Value)
    Check1.ForeColor = kl2
    Check2.ForeColor = kl2
    Check3.ForeColor = kl2
    Option1.ForeColor = kl2
    Option2.ForeColor = kl2
    Option3.ForeColor = kl2
    List1.ForeColor = kl2
    Text1.ForeColor = kl2
    Frame1.ForeColor = kl2
    Frame2.ForeColor = kl2
    Frame3.ForeColor = kl2
    Frame4.ForeColor = kl2
    Frame5.ForeColor = kl2
    Label1.ForeColor = kl2
    Label2.ForeColor = kl2
End Sub
Private Sub HScroll4_Change()
    kl2 = RGB(HScroll6.Value, HScroll5.Value, HScroll4.Value)
    Check1.ForeColor = kl2
    Check2.ForeColor = kl2
    Check3.ForeColor = kl2
    Option1.ForeColor = kl2
    Option2.ForeColor = kl2
    Option3.ForeColor = kl2
    List1.ForeColor = kl2
    Text1.ForeColor = kl2
    Frame1.ForeColor = kl2
    Frame2.ForeColor = kl2
    Frame3.ForeColor = kl2
    Frame4.ForeColor = kl2
    Frame5.ForeColor = kl2
    Label1.ForeColor = kl2
    Label2.ForeColor = kl2
End Sub

Private Sub HScroll6_Scroll()
    kl2 = RGB(HScroll6.Value, HScroll5.Value, HScroll4.Value)
    Check1.ForeColor = kl2
    Check2.ForeColor = kl2
    Check3.ForeColor = kl2
    Option1.ForeColor = kl2
    Option2.ForeColor = kl2
    Option3.ForeColor = kl2
    List1.ForeColor = kl2
    Text1.ForeColor = kl2
    Frame1.ForeColor = kl2
    Frame2.ForeColor = kl2
    Frame3.ForeColor = kl2
    Frame4.ForeColor = kl2
    Frame5.ForeColor = kl2
    Label1.ForeColor = kl2
    Label2.ForeColor = kl2
End Sub
Private Sub HScroll5_Scroll()
    kl2 = RGB(HScroll6.Value, HScroll5.Value, HScroll4.Value)
    Check1.ForeColor = kl2
    Check2.ForeColor = kl2
    Check3.ForeColor = kl2
    Option1.ForeColor = kl2
    Option2.ForeColor = kl2
    Option3.ForeColor = kl2
    List1.ForeColor = kl2
    Text1.ForeColor = kl2
    Frame1.ForeColor = kl2
    Frame2.ForeColor = kl2
    Frame3.ForeColor = kl2
    Frame4.ForeColor = kl2
    Frame5.ForeColor = kl2
    Label1.ForeColor = kl2
    Label2.ForeColor = kl2
End Sub
Private Sub HScroll4_Scroll()
    kl2 = RGB(HScroll6.Value, HScroll5.Value, HScroll4.Value)
    Check1.ForeColor = kl2
    Check2.ForeColor = kl2
    Check3.ForeColor = kl2
    Option1.ForeColor = kl2
    Option2.ForeColor = kl2
    Option3.ForeColor = kl2
    List1.ForeColor = kl2
    Text1.ForeColor = kl2
    Frame1.ForeColor = kl2
    Frame2.ForeColor = kl2
    Frame3.ForeColor = kl2
    Frame4.ForeColor = kl2
    Frame5.ForeColor = kl2
    Label1.ForeColor = kl2
    Label2.ForeColor = kl2
End Sub


Private Sub List1_Click()
    Check1.Font = List1.List(List1.ListIndex)
    Check2.Font = List1.List(List1.ListIndex)
    Check3.Font = List1.List(List1.ListIndex)
    Option1.Font = List1.List(List1.ListIndex)
    Option2.Font = List1.List(List1.ListIndex)
    Option3.Font = List1.List(List1.ListIndex)
    Frame1.Font = List1.List(List1.ListIndex)
    Frame2.Font = List1.List(List1.ListIndex)
    Frame3.Font = List1.List(List1.ListIndex)
    Frame4.Font = List1.List(List1.ListIndex)
    Label1.Font = List1.List(List1.ListIndex)
    Text1.Font = List1.List(List1.ListIndex)
End Sub

