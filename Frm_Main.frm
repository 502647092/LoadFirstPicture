VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Frm_Main 
   Caption         =   "导出首图"
   ClientHeight    =   9390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17160
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   17160
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton alibaba 
      Caption         =   "1688"
      Height          =   300
      Left            =   16530
      TabIndex        =   19
      Top             =   420
      Width           =   600
   End
   Begin VB.CommandButton oa 
      Caption         =   "OA"
      Height          =   300
      Left            =   16530
      TabIndex        =   18
      Top             =   75
      Width           =   600
   End
   Begin VB.TextBox itempicurl 
      Height          =   300
      Index           =   2
      Left            =   12810
      TabIndex        =   15
      Top             =   840
      Width           =   4320
   End
   Begin VB.TextBox itemname 
      Height          =   300
      Index           =   2
      Left            =   9570
      TabIndex        =   14
      Top             =   840
      Width           =   2370
   End
   Begin VB.TextBox urlT 
      Height          =   270
      Index           =   2
      Left            =   945
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   405
      Width           =   7710
   End
   Begin VB.TextBox urlT 
      Height          =   270
      Index           =   1
      Left            =   8760
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   105
      Width           =   7710
   End
   Begin VB.TextBox urlT 
      Height          =   270
      Index           =   0
      Left            =   945
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   105
      Width           =   7710
   End
   Begin 导出商品首图.TzDownload dl 
      Height          =   195
      Index           =   1
      Left            =   30
      Top             =   1185
      Width           =   17070
      _ExtentX        =   30110
      _ExtentY        =   344
      ForeColor       =   33023
   End
   Begin VB.TextBox folder 
      Height          =   300
      Left            =   9795
      TabIndex        =   6
      Top             =   420
      Width           =   6675
   End
   Begin VB.TextBox itemname 
      Height          =   300
      Index           =   1
      Left            =   945
      TabIndex        =   4
      Top             =   840
      Width           =   2370
   End
   Begin VB.TextBox itempicurl 
      Height          =   300
      Index           =   1
      Left            =   4200
      TabIndex        =   2
      Top             =   840
      Width           =   4440
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   7665
      Index           =   0
      Left            =   45
      TabIndex        =   3
      Top             =   1695
      Width           =   8535
      ExtentX         =   15055
      ExtentY         =   13520
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   7665
      Index           =   1
      Left            =   8625
      TabIndex        =   8
      Top             =   1695
      Width           =   8490
      ExtentX         =   14975
      ExtentY         =   13520
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin 导出商品首图.TzDownload dl 
      Height          =   195
      Index           =   2
      Left            =   30
      Top             =   1440
      Width           =   17070
      _ExtentX        =   30110
      _ExtentY        =   344
      ForeColor       =   33023
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   7665
      Index           =   2
      Left            =   8625
      TabIndex        =   13
      Top             =   1710
      Width           =   8490
      ExtentX         =   14975
      ExtentY         =   13520
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "首图链接:"
      Height          =   180
      Left            =   11985
      TabIndex        =   17
      Top             =   900
      Width           =   810
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "商品名称:"
      Height          =   180
      Left            =   8760
      TabIndex        =   16
      Top             =   900
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "网页链接:"
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   450
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "文件夹名称:"
      Height          =   180
      Left            =   8760
      TabIndex        =   7
      Top             =   450
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "商品名称:"
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   900
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "首图链接:"
      Height          =   180
      Left            =   3375
      TabIndex        =   1
      Top             =   900
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "网页链接:"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   810
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strUrl As String

Private Sub alibaba_Click()
    web(0).Navigate2 "http://work.1688.com/home/page/index.htm#nav/home"
End Sub

Private Sub dl_OnFinished(index As Integer, ByVal Result As Boolean)
    dl(index).Tag = True
End Sub

Private Sub dl_OnStart(index As Integer)
    dl(index).Tag = False
End Sub

Private Sub Form_Load()
    web(0).Navigate2 "http://192.168.0.8:83/"
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    Dim lefthg
    lefthg = Me.Height - web(0).Top
    
    web(0).Width = Me.Width - 50
    web(0).Height = lefthg / 3 * 2 - 350
    web(0).Top = 1700
    web(0).Left = 10
    
    web(1).Width = Me.Width - 50
    web(1).Top = 1700 + web(0).Height + 20
    web(1).Height = lefthg / 3 - 350
    web(1).Left = 10
    
    web(2).Width = Me.Width - 50
    web(2).Top = 1700 + web(0).Height + 20
    web(2).Height = lefthg / 3 - 350
    web(2).Left = 10
    
    dl(1).Left = 10
    dl(1).Width = Me.Width - 20
    
    dl(2).Left = 10
    dl(2).Width = Me.Width - 20
End Sub

Private Sub getfp(webb As WebBrowser)
    On Error Resume Next
    Dim i, j, vDoc
    Dim ix As Long
    ix = webb.index
    Set vDoc = webb.Document
    itemname(ix) = resetfilename(vDoc.getelementsbytagname("input")("subject").Value)
    ERR.clear
    itempicurl(ix) = vDoc.getelementsbytagname("input")("pictureUrl").Value
    If ERR <> 0 Then
        itempicurl(ix) = vDoc.getelementsbytagname("input")("pictureUrl")(0).Value
    End If
    
    If folder = "" Then folder = InputBox("请输入 日期-首图-公司名称-提单人名称!", , Format(Now, "m.d") & "-首图-公司名称-提单人名称")
    For i = dl.LBound To dl.UBound
        If dl(i).Tag Then dl(i).FileDownload itempicurl(ix), App.Path & "\" & folder.Text & "\" & itemname(ix).Text & ".jpg": dl(i).Tag = False: Exit For
    Next
End Sub

Private Function resetfilename(ByVal name As String) As String
    name = clear(name, "/")
    name = clear(name, "\")
    name = clear(name, "*")
    name = clear(name, "?")
    name = clear(name, "<")
    name = clear(name, ">")
    resetfilename = name
End Function

Private Function clear(name As String, p As String) As String
    clear = Replace(name, p, "")
End Function

Private Sub Label1_Click()
    Dim i
    For i = web.LBound To web.UBound
        web(i).Stop
    Next
End Sub

Private Sub oa_Click()
    web(0).Navigate2 "http://192.168.0.8:83/"
End Sub

Private Sub urlT_Click(index As Integer)
    urlT(index).SelStart = 0
    urlT(index).SelLength = Len(urlT(index).Text)
End Sub

Private Sub urlT_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then web(index).Navigate2 urlT(index).Text
End Sub

Private Sub web_BeforeNavigate2(index As Integer, ByVal pDisp As Object, url As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    If url <> "http:///" And url <> "" And url <> "about:blank" Then urlT(index) = url
End Sub

Private Sub web_DocumentComplete(index As Integer, ByVal pDisp As Object, url As Variant)
    If InStr(1, url, "operator=edit") Then Call getfp(web(index))
End Sub

Private Sub web_DownloadBegin(index As Integer)
    web(index).Tag = False
End Sub

'Private Sub web_DownloadBegin(index As Integer)
'    web(index).Silent = True
'End Sub

Private Sub web_DownloadComplete(index As Integer)
    web(index).Silent = True
    web(index).Tag = True
    showweb (index)
End Sub

Private Sub web_NewWindow2(index As Integer, ppDisp As Object, Cancel As Boolean)
    Dim i
    For i = 1 To web.UBound
        If web(i).Tag Then Set ppDisp = web(i).Object: showweb (i): Exit For
    Next
End Sub

Private Sub showweb(index As Long)
    Dim i As Long
    For i = 1 To web.UBound
        web(i).Visible = False
    Next
    web(index).Visible = True
End Sub
