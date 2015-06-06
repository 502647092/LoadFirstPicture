VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Frm_Main 
   Caption         =   "导出首图"
   ClientHeight    =   9390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15750
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   15750
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton getitem 
      Caption         =   "Item"
      Height          =   375
      Left            =   10440
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton getmain 
      Caption         =   "Main"
      Height          =   375
      Left            =   9000
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox mainurl 
      Height          =   300
      Left            =   1125
      TabIndex        =   9
      Text            =   "I:\工作空间\VB工程文件\Git工程文件\首图导出工具\fp.html"
      Top             =   90
      Width           =   7575
   End
   Begin VB.TextBox itemurl 
      Height          =   300
      Left            =   8760
      TabIndex        =   8
      Text            =   "I:\工作空间\VB工程文件\Git工程文件\首图导出工具\fp.html"
      Top             =   90
      Width           =   6975
   End
   Begin SHDocVwCtl.WebBrowser item 
      Height          =   7650
      Left            =   8175
      TabIndex        =   7
      Top             =   1695
      Width           =   7530
      ExtentX         =   13282
      ExtentY         =   13494
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
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
   Begin VB.TextBox folder 
      Height          =   300
      Left            =   1125
      TabIndex        =   5
      Text            =   "首图"
      Top             =   990
      Width           =   7575
   End
   Begin 导出商品首图.TzDownload dl 
      Height          =   240
      Left            =   120
      Top             =   1395
      Width           =   15555
      _ExtentX        =   27437
      _ExtentY        =   423
      ForeColor       =   33023
   End
   Begin VB.TextBox itempicurl 
      Height          =   300
      Left            =   1125
      TabIndex        =   4
      Top             =   690
      Width           =   7575
   End
   Begin VB.TextBox itemname 
      Height          =   300
      Left            =   1125
      TabIndex        =   3
      Top             =   390
      Width           =   7575
   End
   Begin SHDocVwCtl.WebBrowser main 
      Height          =   7995
      Left            =   45
      TabIndex        =   10
      Top             =   1695
      Width           =   8085
      ExtentX         =   14261
      ExtentY         =   14102
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "文件夹名称:"
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "首图链接:"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   765
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "商品名称:"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   465
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "商品链接:"
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
Private Sub dl_OnFinished(ByVal Result As Boolean)
    If Result Then
        
    Else
    
    End If
End Sub

Private Sub goto_Click()
    main.Navigate2 mainurl.Text
End Sub

Private Sub Form_Load()
 main.Navigate2 "http://192.168.0.8:83/"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    main.Width = Me.Width - 50
    item.Width = Me.Width - 50
    
    Dim lefthg
    lefthg = Me.Height - main.Top
    
    main.Height = lefthg / 2 - 350
    item.Height = lefthg / 2 - 350
    
    main.Top = 1700
    item.Top = 1700 + main.Height + 20
    
    main.Left = 10
    item.Left = 10
    
    dl.Left = 10
    dl.Width = Me.Width - 20
    
End Sub

Private Sub getmain_Click()
    Call getfp(main)
End Sub

Private Sub getitem_Click()
    Call getfp(item)
End Sub

Private Sub getfp(web As WebBrowser)
    On Error Resume Next
    Dim i, j, vDoc
    Set vDoc = web.Document
    itemname = resetfilename(vDoc.getelementsbytagname("input")("subject").Value)
    itempicurl = vDoc.getelementsbytagname("input")("pictureUrl")(0).Value
    dl.FileDownload itempicurl, App.Path & "\" & folder & "\" & itemname & ".jpg"
End Sub

Private Sub item_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    If InStr(1, URL, "operator=edit") Then Call getfp(item)
End Sub

Private Sub main_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    If InStr(1, URL, "operator=edit") Then Call getfp(main)
End Sub

Private Sub item_NewWindow2(ppDisp As Object, Cancel As Boolean)
    Set ppDisp = main.Object
End Sub

Private Sub main_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    If URL <> "http:///" And URL <> "" And URL <> "about:blank" Then mainurl = URL
End Sub

Private Sub item_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    If URL <> "http:///" And URL <> "" And URL <> "about:blank" Then itemurl = URL
End Sub

Private Sub main_NewWindow2(ppDisp As Object, Cancel As Boolean)
    Set ppDisp = item.Object
    'Cancel = True
    'item.Navigate strUrl
End Sub

Private Sub main_DownloadBegin()
    main.Silent = True
End Sub

Private Sub main_DownloadComplete()
    main.Silent = True
End Sub

Private Sub item_DownloadBegin()
    item.Silent = True
End Sub

Private Sub item_DownloadComplete()
    item.Silent = True
End Sub

Private Sub itemurl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then item.Navigate2 itemurl.Text
End Sub

Private Sub mainurl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then main.Navigate2 mainurl.Text
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

