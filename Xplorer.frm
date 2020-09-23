VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Xplorer"
   ClientHeight    =   4470
   ClientLeft      =   1200
   ClientTop       =   2880
   ClientWidth     =   8070
   Icon            =   "Xplorer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   8070
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   300
      Left            =   6960
      TabIndex        =   11
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ComboBox cmbSearch 
      Height          =   315
      Left            =   5280
      TabIndex        =   10
      Text            =   "Input your search phrase in here"
      ToolTipText     =   "Input the search phrase here"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Text            =   "Bookmarks"
      ToolTipText     =   "Browse you bookmarks here"
      Top             =   1200
      Width           =   5175
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   7
      ToolTipText     =   "Type in the URL's here"
      Top             =   480
      Width           =   5175
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go!"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3600
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1680
      Picture         =   "Xplorer.frx":030A
      TabIndex        =   5
      Top             =   840
      Width           =   1935
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4200
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   "Forward"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   4575
      ExtentX         =   8070
      ExtentY         =   2990
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label lblStatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5280
      TabIndex        =   8
      Top             =   120
      Width           =   4695
      WordWrap        =   -1  'True
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Home 
         Caption         =   "Homepage"
      End
      Begin VB.Menu NewWindow 
         Caption         =   "New Window"
         Shortcut        =   ^N
      End
      Begin VB.Menu Save 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu Properties 
         Caption         =   "Properties"
      End
      Begin VB.Menu PageSetp 
         Caption         =   "Page Setup"
      End
      Begin VB.Menu Print 
         Caption         =   "Print"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edit"
      Begin VB.Menu Cut 
         Caption         =   "Cut"
      End
      Begin VB.Menu Copy 
         Caption         =   "Copy"
      End
      Begin VB.Menu Paste 
         Caption         =   "Paste"
      End
      Begin VB.Menu SelectAll 
         Caption         =   "Select All"
      End
   End
   Begin VB.Menu Bookmarks 
      Caption         =   "Bookmarks"
      Begin VB.Menu ClearBkmarks 
         Caption         =   "Clear Bookmarks"
      End
      Begin VB.Menu Add 
         Caption         =   "Add a Bookmark"
         Shortcut        =   +{INSERT}
      End
   End
   Begin VB.Menu Options 
      Caption         =   "Options"
      Begin VB.Menu Popups 
         Caption         =   "Disable Popups"
         Shortcut        =   ^P
      End
      Begin VB.Menu Clear 
         Caption         =   "Clear History"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu About 
         Caption         =   "About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public AllowPopups As Boolean
Public NumberOfTimesClicked As Integer
Public Index As Integer
Public State As Integer
Dim mbDontNavigateNow As Boolean
    
Private Sub About_Click()
    Form2.Visible = True
End Sub

Private Sub Add_Click()
    AddBookmarks (App.Path & "\bookmarks.txt")
End Sub

Private Sub Clear_Click()
    On Error Resume Next
    Combo1.Clear
    Combo1.Text = WebBrowser1.LocationURL
    Dim i As Integer
    Dim a As String
    Open App.Path & "\history.txt" For Output As #1
    For i = 0 To Combo1.ListCount - 1
    Write #1, ""
    Next i
    Close #1
End Sub

Private Sub ClearBkmarks_Click()
    On Error Resume Next
    Combo2.Clear
    Dim i As Integer
    Dim a As String
    Open App.Path & "\bookmarks.txt" For Output As #1
    For i = 0 To Combo2.ListCount - 1
    Write #1, ""
    Next i
    Close #1
    Combo2.Text = "Bookmarks"
End Sub

Private Sub cmbSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    cmdSearch.Default = True
End Sub

Private Sub cmdStop_Click()
    WebBrowser1.Stop
    ProgressBar1.Visible = False
    WebBrowser1.Height = Form1.ScaleHeight - 1520
    lblStatus.Caption = "Loading stopped"
    Combo1.SetFocus
End Sub

Private Sub cmdGo_Click()
    WebBrowser1.Navigate (Combo1.Text)
End Sub

Private Sub cmdRefresh_Click()
    On Error Resume Next
    WebBrowser1.Refresh
End Sub

Private Sub Combo2_Change()
    WebBrowser1.Navigate Combo2.SelText
End Sub

Private Sub cmdSearch_Click()
    On Error Resume Next
    WebBrowser1.Navigate ("http://search.dogpile.com/texis/search?q=" & cmbSearch.Text & "&geo=no&refer=dp-search&fs=web")
    cmbSearch.AddItem (cmbSearch.Text)
    cmbSearch.SetFocus
End Sub

Private Sub Copy_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub Cut_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_CUT, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub Exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    State = 1
    WebBrowser1.GoHome
    NumberOfTimesClicked = 0
    AllowPopups = True
    LoadBookmarks
    Dim a As String
    On Error Resume Next
    Open App.Path & "\history.txt" For Input As #1
    Do
        Input #1, a
        If a <> "" Then
        Combo1.AddItem a
    End If
    Loop Until EOF(1)
    Close #1
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Form1.WindowState <> 1 Then
        WebBrowser1.Width = Form1.ScaleWidth
        WebBrowser1.Height = Form1.ScaleHeight - 1525
        ProgressBar1.Width = Form1.ScaleWidth
        ProgressBar1.Top = Form1.ScaleHeight - 250
        lblStatus.Width = Form1.ScaleWidth - Combo1.Width
        cmbSearch.Width = Form1.Width - Combo1.Width - cmdSearch.Width - 300
        cmdSearch.Left = Form1.ScaleWidth - cmdSearch.Width
    End If
End Sub

Private Sub cmdBack_Click()
    On Error Resume Next
    WebBrowser1.GoBack
End Sub

Private Sub Form_Terminate()
    Unload Me
End Sub

Private Sub LoadBookmarks()
    Dim a As String
    On Error Resume Next
    Open App.Path & "\bookmarks.txt" For Input As #1
    Do
    Input #1, a
    If a <> "" Then
    Combo2.AddItem a
    End If
    Loop Until EOF(1)
    Close #1
End Sub

Private Sub Home_Click()
    WebBrowser1.GoHome
End Sub

Private Sub NewWindow_Click()
    On Error Resume Next
    Static lDocumentCount As Long
    Dim frmD As Form
    lDocumentCount = lDocumentCount + 1
    Set frmD = New Form1
    frmD.Show
    frmD.SetFocus
End Sub

Private Sub AddBookmarks(filename As String)
    Combo2.AddItem WebBrowser1.LocationURL
    Dim i As Integer
    Dim a As String
    Dim URL As String
    Open App.Path & "\bookmarks.txt" For Output As #1
    For i = 0 To Combo1.ListCount + 1
    Write #1, Combo2.List(i)
    Next i
    Close #1
End Sub
Private Sub PageSetp_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub Paste_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub Popups_Click()
    If Popups.Checked = False Then
        Popups.Checked = True
        AllowPopups = False
    ElseIf Popups.Checked = True Then
        Popups.Checked = False
        AllowPopups = True
    End If
End Sub

Private Sub Print_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub Properties_Click()
    WebBrowser1.ExecWB OLECMDID_PROPERTIES, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub Save_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub SelectAll_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub WebBrowser1_DownloadBegin()
    ProgressBar1.Visible = True
    WebBrowser1.Height = Form1.ScaleHeight - 1855
    ProgressBar1.Max = 1
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    Combo1.Text = WebBrowser1.LocationURL
    Form1.Caption = "Xplorer - " + WebBrowser1.LocationName
    ProgressBar1.Visible = False
    WebBrowser1.Height = Form1.ScaleHeight - 1520
    Combo1.AddItem Combo1.Text
    Dim i As Integer
    Dim a As String
    Open App.Path & "\history.txt" For Output As #1
    For i = 0 To Combo1.ListCount - 1
    Write #1, Combo1.List(i)
    Next i
    Close #1
    cmdGo.Default = True
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
    If AllowPopups = True Then
        Cancel = False
        DoEvents
    ElseIf AllowPopups = False Then
        Cancel = True
    End If
End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    If ProgressMax >= 0 And Progress > 0 And Progress <= ProgressMax Then
        ProgressBar1.Value = Progress / ProgressMax
    End If
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
    lblStatus.Caption = Text
End Sub

Private Sub cmdForward_Click()
    On Error Resume Next
    WebBrowser1.GoForward
End Sub
