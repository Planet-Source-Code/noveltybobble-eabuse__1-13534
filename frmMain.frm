VERSION 5.00
Object = "{2F7793D6-84FE-11D4-84A3-00508BF5022F}#2.0#0"; "MSSlider.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmMain 
   Caption         =   "EAbuse Beta 2.0"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4875
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btn_DoIt 
      Caption         =   "Create Abuse"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "frmMain.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Width           =   4575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4815
      Begin MSMAPI.MAPIMessages MAPIMessages1 
         Left            =   3960
         Top             =   1560
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         AddressEditFieldCount=   1
         AddressModifiable=   0   'False
         AddressResolveUI=   0   'False
         FetchSorted     =   0   'False
         FetchUnreadOnly =   0   'False
      End
      Begin MSMAPI.MAPISession MAPISession1 
         Left            =   3360
         Top             =   1560
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DownloadMail    =   -1  'True
         LogonUI         =   -1  'True
         NewSession      =   0   'False
      End
      Begin VB.TextBox txtWordCount 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Text            =   "5"
         Top             =   1440
         Width           =   1095
      End
      Begin VBMSSlider.MSSlider Slider1 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Fuck off"
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
      End
      Begin VB.Label Label4 
         Caption         =   "Word Count"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Mild"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Outrageous"
         Height          =   255
         Left            =   1800
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Complete Cunt"
         Height          =   255
         Left            =   3480
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Abusive Result"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   4815
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   2040
         TabIndex        =   11
         Text            =   "John"
         Top             =   3240
         Width           =   2655
      End
      Begin VB.CommandButton btnMail 
         Caption         =   "Mail Abuse To:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox txtResult 
         Height          =   1935
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1200
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private colHigh As New Collection
Private colMedium As New Collection
Private colLow As New Collection
Private lHighCount As Long
Private lMediumCount As Long
Private lLowCount As Long

Private Sub btn_DoIt_Click()
    Dim lLowIndex As Long
    Dim lMediumIndex As Long
    Dim lHighIndex As Long
    Dim lSeed As Long
    Dim lLoop As Long
    Dim lWhichInsult As Long
    Dim Entry As CAbuseEntry
    
    lWhichInsult = 3
    txtResult.Text = ""
    
    For lLoop = 1 To CLng(txtWordCount.Text)
        lLowIndex = Rand(1, lLowCount)
        lMediumIndex = Rand(1, lMediumIndex)
        lHighIndex = Rand(1, lHighCount)
        
        If lWhichInsult > Slider1.Value Then lWhichInsult = 1
        
        Select Case lWhichInsult
        Case 1
            Set Entry = colLow.Item(lLowIndex)
            txtResult = txtResult & Entry.strAbuse & " "
        Case 2
            Set Entry = colMedium.Item(lMediumIndex)
            txtResult = txtResult & Entry.strAbuse & " "
        Case 3
            Set Entry = colHigh.Item(lHighIndex)
            txtResult = txtResult & Entry.strAbuse & " "
        End Select
        
        lWhichInsult = lWhichInsult + 1
    Next
        


End Sub

Private Sub btnMail_Click()
    MailResults txtEmail
End Sub

Private Sub Form_Load()
    Slider1.Min = 1
    Slider1.Max = 3
    Slider1.LowColor = vbBlue
    Slider1.HiColor = vbRed
    Slider1.ColorShift = True
    Slider1.ColInit
    Slider1.GripperPic = App.Path & "\SliderButton.bmp"
    
    LoadFiles
    Randomize (Second(Now))
End Sub




Private Sub txtWordCount_Validate(Cancel As Boolean)
    If Not IsNumeric(txtWordCount) Then Cancel = True
End Sub


Private Sub LoadFiles()
    Dim File As New CFile
    Dim s As String
    Dim bResult As Boolean
    Dim Entry As CAbuseEntry
    
    'On Error GoTo ErrorHandler
    
    '//
    '//     Retrieve Low
    '//
    bResult = File.SourceFile(App.Path & "\abuse-low.twat")
    If Not bResult Then
        txtResult = "Low abuse file is missing"
    Else
        Do While File.GetString(s) = True
            Set Entry = New CAbuseEntry
            Entry.strAbuse = s
            lLowCount = lLowCount + 1
            colLow.Add Entry, CStr(lLowCount)
        Loop
    End If
    
    
    '//
    '//     Retrieve Medium
    '//
    bResult = File.SourceFile(App.Path & "\abuse-medium.twat")
    If Not bResult Then
        txtResult = txtResult & vbCrLf & "Medium abuse file is missing"
    Else
        Do While File.GetString(s) = True
            Set Entry = New CAbuseEntry
            Entry.strAbuse = s
            lMediumCount = lMediumCount + 1
            colMedium.Add Entry, CStr(lMediumCount)
        Loop
    End If
        
    '//
    '//     Retrieve High
    '//
    bResult = File.SourceFile(App.Path & "\abuse-High.twat")
    If Not bResult Then
        txtResult = txtResult & vbCrLf & "High abuse file is missing"
    Else
        Do While File.GetString(s) = True
            Set Entry = New CAbuseEntry
            Entry.strAbuse = s
            lHighCount = lHighCount + 1
            colHigh.Add Entry, CStr(lHighCount)
        Loop
    End If
    
    Exit Sub
    
ErrorHandler:
    txtResult = Err.Description
End Sub


Public Function Rand(ByVal Low As Long, _
                     ByVal High As Long) As Long
' The random number equation
  Rand = Int((High - Low + 1) * Rnd) + Low
End Function






Private Sub MailResults(strAddress As String)
    On Local Error GoTo ReportError
    
    MAPISession1.SignOn
    MAPIMessages1.SessionID = MAPISession1.SessionID
    MAPIMessages1.MsgIndex = -1
    
    
    With MAPIMessages1
        .AddressResolveUI = False
        .MsgNoteText = txtResult.Text
        .MsgSubject = ""
        .RecipAddress = strAddress
        .RecipDisplayName = strAddress
        .ResolveName
        .Send
    End With
    
    MAPISession1.SignOff
    Exit Sub
ReportError:
    txtResult = txtResult & vbCrLf & Err.Description
End Sub



