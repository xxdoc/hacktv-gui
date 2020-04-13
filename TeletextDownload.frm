VERSION 5.00
Begin VB.Form TeletextDownload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Download teletext"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin hacktv_gui.NetGrab NetGrab1 
      Left            =   3960
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.CommandButton DownloadSpark 
      Caption         =   "SPARK"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton DownloadCancel 
      Caption         =   "Close"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton DownloadTeefax 
      Caption         =   "Teefax"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Status 
      Caption         =   "Ready to download"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Please choose a teletext service to download."
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label4 
      Caption         =   "Close without downloading"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Download SPARK from TVARK"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Download Teefax"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "TeletextDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FilePath As String
Dim TeletextPath As String
Dim DownloadURL As String
Dim strText As String, strData() As String
Dim lngA As Long
Dim totalcounter As Integer
Dim progresscounter As Integer
Dim DownloadInProgress As Boolean
Dim DownloadCancelled As Boolean

Private Sub Form_Load()
' Set the form's icon to the same icon used in the main form
    Me.Icon = MainForm.Icon
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If DownloadInProgress = True Then
        DownloadCancelled = True
        DownloadTeefax.Enabled = True
        DownloadSpark.Enabled = True
        DownloadCancel.Caption = "Close"
        Status.Caption = "Cancelled"
        MousePointer = 0
        Cancel = True
    End If
End Sub

Private Sub DownloadCancel_Click()
    Unload Me
End Sub

Private Sub DownloadSpark_Click()
' Set variables
    TeletextPath = Environ$("temp") & Chr(92) & "spark"
    FilePath = Environ$("temp") & "\spark.html"
    DownloadURL = "https://raw.githubusercontent.com/ZXGuesser/spark-teletext/master/"
' Prevent the form from hanging
    DoEvents
' Disable download buttons
    DownloadTeefax.Enabled = False
    DownloadSpark.Enabled = False
    DownloadCancel.Caption = "Cancel"
' Reset file counters to zero (used for the status bar)
    progresscounter = 0
    totalcounter = 0
    Status.Caption = "Clearing cache..."
    DeleteUrlCacheEntry DownloadURL
    Status.Caption = "Downloading index from Github..."
    NetGrab1.DownloadStart "https://github.com/ZXGuesser/spark-teletext", vbAsyncReadSynchronousDownload
    If ReadHTMLFile = False Then Exit Sub
    If DownloadCancelled = False Then
        MainForm.teletext_source.Text = TeletextPath
        Unload Me
    End If
End Sub

Private Sub DownloadTeefax_Click()
' Set variables
    TeletextPath = Environ$("temp") & Chr(92) & "teefax"
    FilePath = Environ$("temp") & "\teefax.html"
    DownloadURL = "http://teastop.plus.com/svn/teletext/"
' Prevent the form from hanging
    DoEvents
' Disable download buttons
    DownloadTeefax.Enabled = False
    DownloadSpark.Enabled = False
    DownloadCancel.Caption = "Cancel"
' Reset file counters to zero (used for the status bar)
    progresscounter = 0
    totalcounter = 0
    Status.Caption = "Clearing cache..."
    DeleteUrlCacheEntry DownloadURL
    Status.Caption = "Downloading Teefax index..."
    NetGrab1.DownloadStart DownloadURL, vbAsyncReadSynchronousDownload
    If ReadHTMLFile = False Then Exit Sub
    If DownloadCancelled = False Then
        MainForm.teletext_source.Text = TeletextPath
        Unload Me
    End If
End Sub

Private Function ReadHTMLFile() As Boolean
    MousePointer = 13
    ReadHTMLFile = True
    DownloadInProgress = False
    DownloadCancelled = False
    Dim strText As String, strData() As String
    Dim lngA As Long
    Dim iFileNo As Integer
    iFileNo = FreeFile
    If FolderExists(TeletextPath) Then
        If Not PathIsDirectoryEmpty(TeletextPath) = "1" Then Kill (TeletextPath & Chr(92) & "*.*")
    Else
        MkDir (TeletextPath)
    End If
    If DoesFileExist(FilePath) = False Then
        MsgBox "Index page did not download. Please ensure that you are connected to the internet and try again.", vbExclamation, App.Title
        Status.Caption = "Cancelled"
        DownloadTeefax.Enabled = True
        DownloadSpark.Enabled = True
        DownloadCancel.Caption = "Close"
        ReadHTMLFile = False
        MousePointer = 0
        Exit Function
    End If
' open data from file
    Open FilePath For Input As #iFileNo
    strText = Input(LOF(1), #iFileNo)
    Close #iFileNo
' Parse the file twice, the first time we run a simple counter to check
' how many files we have to download
    For lngA = 0 To Between(strText, ".tti" & Chr(34) & ">", "</a>", strData) - 1
        totalcounter = totalcounter + 1
        Next lngA
        Debug.Print (totalcounter & " files to download")
' if we get more than none...
        Status.Caption = "Parsing index file..."
        DownloadInProgress = True
        For lngA = 0 To Between(strText, ".tti" & Chr(34) & ">", "</a>", strData) - 1
' Download everything we found
            If DownloadCancelled = True Then Exit For
            DoEvents ' Prevent the form from hanging
            progresscounter = progresscounter + 1
            Status.Caption = "Downloading page " & strData(lngA) & Chr(32) & "(" & progresscounter & " of " & totalcounter & ")"
            FilePath = TeletextPath & Chr(92) & strData(lngA)
            DeleteUrlCacheEntry DownloadURL & strData(lngA)
            NetGrab1.DownloadStart DownloadURL & strData(lngA), vbAsyncReadSynchronousDownload
            Next lngA
    If DownloadCancelled = False Then
        Status.Caption = "Done"
    Else
        Status.Caption = "Cancelled"
    End If
    MousePointer = 0
    DownloadInProgress = False
End Function
        
Private Sub NetGrab1_DownloadComplete(ByVal nBytes As Long)
    Call NetGrab1.SaveAs(FilePath)
End Sub
