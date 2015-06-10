VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Frm_Main 
   Caption         =   "导出首图"
   ClientHeight    =   9390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19035
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   19035
   StartUpPosition =   3  '窗口缺省
   Begin 导出商品首图.TzListBox UName 
      Height          =   1170
      Left            =   8835
      TabIndex        =   28
      Top             =   4440
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   2064
   End
   Begin 导出商品首图.TzListBox SName 
      Height          =   1335
      Left            =   8835
      TabIndex        =   27
      Top             =   3135
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   2355
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   915
      Index           =   1
      Left            =   8580
      TabIndex        =   8
      Top             =   1335
      Width           =   1155
      ExtentX         =   2037
      ExtentY         =   1614
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
      Height          =   915
      Index           =   2
      Left            =   9750
      TabIndex        =   26
      Top             =   1350
      Width           =   1155
      ExtentX         =   2037
      ExtentY         =   1614
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
   Begin VB.CommandButton oa 
      Caption         =   "OA"
      Height          =   300
      Left            =   16425
      TabIndex        =   23
      Top             =   60
      Width           =   600
   End
   Begin VB.CommandButton alibaba 
      Caption         =   "1688"
      Height          =   300
      Left            =   16425
      TabIndex        =   22
      Top             =   360
      Width           =   600
   End
   Begin VB.CommandButton manager 
      Caption         =   "商品"
      Height          =   300
      Left            =   16425
      TabIndex        =   21
      Top             =   660
      Width           =   600
   End
   Begin VB.CommandButton pic 
      Caption         =   "图片"
      Height          =   300
      Left            =   16425
      TabIndex        =   20
      Top             =   960
      Width           =   600
   End
   Begin VB.ListBox List1 
      Height          =   7620
      Left            =   16830
      TabIndex        =   18
      Top             =   1695
      Width           =   1035
   End
   Begin VB.ListBox List2 
      Height          =   780
      Left            =   45
      TabIndex        =   17
      Top             =   9315
      Width           =   17145
   End
   Begin VB.TextBox itempicurl 
      Height          =   300
      Index           =   2
      Left            =   12810
      TabIndex        =   14
      Top             =   960
      Width           =   3600
   End
   Begin VB.TextBox itemname 
      Height          =   300
      Index           =   2
      Left            =   9570
      TabIndex        =   13
      Top             =   960
      Width           =   2370
   End
   Begin VB.TextBox urlT 
      Height          =   270
      Index           =   2
      Left            =   8700
      TabIndex        =   12
      Text            =   "2"
      Top             =   420
      Width           =   7710
   End
   Begin VB.TextBox urlT 
      Height          =   270
      Index           =   1
      Left            =   945
      TabIndex        =   10
      Text            =   "1"
      Top             =   420
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
      Left            =   930
      Top             =   713
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   344
      ForeColor       =   33023
   End
   Begin VB.TextBox folder 
      Height          =   300
      Left            =   9735
      TabIndex        =   6
      Top             =   90
      Width           =   6675
   End
   Begin VB.TextBox itemname 
      Height          =   300
      Index           =   1
      Left            =   1815
      TabIndex        =   4
      Top             =   960
      Width           =   2370
   End
   Begin VB.TextBox itempicurl 
      Height          =   300
      Index           =   1
      Left            =   5070
      TabIndex        =   2
      Top             =   960
      Width           =   3600
   End
   Begin 导出商品首图.TzDownload dl 
      Height          =   195
      Index           =   2
      Left            =   8700
      Top             =   713
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   344
      ForeColor       =   33023
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   7665
      Index           =   0
      Left            =   45
      TabIndex        =   3
      Top             =   1335
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
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "首图信息:"
      Height          =   180
      Left            =   120
      TabIndex        =   25
      Top             =   1020
      Width           =   810
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "下载状态:"
      Height          =   180
      Left            =   120
      TabIndex        =   24
      Top             =   720
      Width           =   810
   End
   Begin VB.Label pages 
      AutoSize        =   -1  'True
      Caption         =   "页数"
      Height          =   180
      Left            =   17310
      TabIndex        =   19
      Top             =   885
      Width           =   360
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "首图链接:"
      Height          =   180
      Left            =   11985
      TabIndex        =   16
      Top             =   1020
      Width           =   810
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "商品名称:"
      Height          =   180
      Left            =   8760
      TabIndex        =   15
      Top             =   1020
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "网页链接:"
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   465
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "文件夹名称:"
      Height          =   180
      Left            =   8700
      TabIndex        =   7
      Top             =   150
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "商品名称:"
      Height          =   180
      Left            =   990
      TabIndex        =   5
      Top             =   1020
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "首图链接:"
      Height          =   180
      Left            =   4245
      TabIndex        =   1
      Top             =   1020
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "主页链接:"
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

Private Sub Form_Load()
    web(0).Navigate2 "http://192.168.0.8:83/"
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    Dim lefthg
    web(0).Top = 1300
    lefthg = Me.Height - web(0).Top
    
    web(0).Width = Me.Width - 50
    web(0).Height = lefthg - 250
    web(0).Left = 10
    Dim i As Long
    For i = 1 To web.UBound
        web(i).Width = Me.Width / 3 * 2 - 50
        web(i).Top = web(0).Top + lefthg / 2
        web(i).Height = lefthg / 2
        web(i).Left = 10
    Next
'    web(1).Width = Me.Width / 3 * 2 - 50
'    web(1).Top = web(0).Top + lefthg / 2
'    web(1).Height = lefthg / 2
'    web(1).Left = 10
'
'    web(2).Width = Me.Width / 3 * 2 - 50
'    web(2).Top = web(0).Top + lefthg / 2
'    web(2).Height = lefthg / 2
'    web(2).Left = 10
    
    SName.Width = web(0).Width / 2
    SName.Top = web(0).Top + lefthg / 2
    SName.Height = lefthg / 3 * 2
    SName.Left = 10
    
    UName.Width = web(0).Width / 2
    UName.Top = web(0).Top + lefthg / 2
    UName.Height = lefthg / 3 * 2
    UName.Left = 10 + SName.Width
    
    
    List1.Left = Me.Width - List1.Width - 350
    List1.Height = lefthg - 350
    List1.Top = web(0).Top
End Sub

Private Sub getfp(webb As WebBrowser)
    On Error Resume Next
    Dim i, J, vDoc
    Dim ix As Long
    ix = webb.index
    Set vDoc = webb.Document
    itemname(ix) = resetfilename(vDoc.getelementsbytagname("input")("subject").Value)
    ERR.Clear
    itempicurl(ix) = vDoc.getelementsbytagname("input")("pictureUrl").Value
    If ERR <> 0 Then
        itempicurl(ix) = vDoc.getelementsbytagname("input")("pictureUrl")(0).Value
    End If
    
    If folder = "" Then folder = InputBox("请输入 日期-首图-公司名称-阿里账号-提单人名称!", , Format(Now, "m.d") & "-首图-公司名称-阿里账号-提单人名称")
    If folder = "" Then folder = Format(Now, "m.d") & "-首图-公司名称-阿里账号-提单人名称"
    If Len(itemname(ix)) = 0 And Len(itempicurl(ix)) = 0 Then Exit Sub
    For i = dl.LBound To dl.UBound
        If dl(i).IsFree Then dl(i).FileDownload itempicurl(ix), App.Path & "\" & folder.Text & "\" & itemname(ix).Text & ".jpg": dl(i).Tag = False: Exit For
    Next
    itemname(ix) = ""
    itempicurl(ix) = ""
End Sub

Private Function resetfilename(ByVal name As String) As String
    name = Clear(name, "/")
    name = Clear(name, "\")
    name = Clear(name, "*")
    name = Clear(name, "?")
    name = Clear(name, "<")
    name = Clear(name, ">")
    resetfilename = name
End Function

Private Function Clear(name As String, p As String) As String
    Clear = Replace(name, p, "")
End Function

Private Sub Label1_Click()
    web(0).Visible = Not web(0).Visible
    showweb (0)
End Sub

Private Sub Label2_Click()
'A A   A   DIV DIV
    Dim vDoc, vTag_2, vTag_1, vTag, vTag1, vTag2, vTXT
    Dim i As Integer
    Set vDoc = web(0).Document
    On Error Resume Next
    For i = 0 To vDoc.All.length - 1
        List1.AddItem vDoc.All(i).TagName
        Set vTag_2 = vDoc.All(i - 2)
        Set vTag_1 = vDoc.All(i - 1)
        Set vTag = vDoc.All(i)
        Set vTag1 = vDoc.All(i + 1)
        Set vTag2 = vDoc.All(i + 2)
        Select Case UCase(vDoc.All(i).TagName)
        Case "A"
            If UCase(vTag_2.TagName) = "A" And _
               UCase(vTag_1.TagName) = "A" And _
               UCase(vTag1.TagName) = "DIV" And _
               UCase(vTag2.TagName) = "DIV" Then
                If vTag.class = "next" Then vTag.Click
            End If
        End Select
    Next
End Sub

Private Sub Label5_Click()
    On Error Resume Next
    Dim i As Long
    For i = web.LBound To web.UBound
        web(i).Stop
        web(i).Tag = True
    Next
End Sub

Private Sub manager_Click()
    web(0).Navigate2 "http://offer.1688.com/offer/manage.htm?show_type=valid&tracelog=work_1_m_orderManage"
End Sub
'http://picman.1688.com/album/album_list.htm?tracelog=work_1_m_albumManage
'http://offer.1688.com/offer/manage.htm?show_type=valid&tracelog=work_1_m_orderManage
'http://login.1688.com/member/signout.htm
Private Sub oa_Click()
    web(0).Navigate2 "http://192.168.0.8:83/"
End Sub

Private Sub pic_Click()
    web(0).Navigate2 "http://picman.1688.com/album/album_list.htm?tracelog=work_1_m_albumManage"
End Sub

Private Sub urlT_DblClick(index As Integer)
    urlT(index).SelStart = 0
    urlT(index).SelLength = Len(urlT(index).Text)
End Sub

Private Sub urlT_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then web(index).Navigate2 urlT(index).Text
End Sub

Private Sub web_BeforeNavigate2(index As Integer, ByVal pDisp As Object, url As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    If url <> "http:///" And url <> "" And url <> "about:blank" Then urlT(index) = url
    'List3.AddItem url
End Sub

Private Sub web_DocumentComplete(index As Integer, ByVal pDisp As Object, url As Variant)
    On Error Resume Next
    If InStr(1, url, "operator=edit") Then Call getfp(web(index))
End Sub

Private Sub web_DownloadBegin(index As Integer)
    web(index).Tag = False
    urlT(index).Enabled = False
    Me.Caption = "Loading..."
End Sub

'Private Sub web_DownloadBegin(index As Integer)
'    web(index).Silent = True
'End Sub

Private Sub web_DownloadComplete(index As Integer)
    Dim target, title, class
    web(index).Silent = True
    web(index).Tag = True
    urlT(index).Enabled = True
    urlT(index).ForeColor = vbBlue
    Me.Caption = "Load Complete"
    showweb (index)
    List1.Clear
    List2.Clear
    Dim vDoc, vTag_2, vTag_1, vTag, vTag1, vTag2, vTXT
    Dim i As Integer
    Set vDoc = web(index).Document
    'On Error Resume Next
    For i = 0 To vDoc.All.length - 1
        List1.AddItem vDoc.All(i).TagName
        On Error Resume Next
        Set vTag_2 = vDoc.All(i - 2)
        Set vTag_1 = vDoc.All(i - 1)
        Set vTag = vDoc.All(i)
        Set vTag1 = vDoc.All(i + 1)
        Set vTag2 = vDoc.All(i + 2)
        Select Case UCase(vDoc.All(i).TagName)
        Case "TD"
        Case "A"
        Dim st
        Dim en
            If UCase(vTag_2.TagName) = "INPUT" And _
               UCase(vTag_1.TagName) = "TD" And _
               UCase(vTag1.TagName) = "IMG" And _
               UCase(vTag2.TagName) = "TD" Then
                If vTag.target = "_blank" Then
                    If SName.AddItemNotSame(vTag.title) Then
                        st = InStr(1, vTag.innerhtml, "data-lazyload-src=""") + Len("data-lazyload-src=""")
                        en = InStr(st, vTag.innerhtml, """")
                        UName.AddItemNotSame Replace(Mid(vTag.innerhtml, st, en - st), ".64x64", "")
                    End If
                End If
            End If
'            If UCase(vTag_2.TagName) = "DIV" And _
'               UCase(vTag_1.TagName) = "DIV" And _
'               UCase(vTag1.TagName) = "SPAN" And _
'               UCase(vTag2.TagName) = "UL" Then
'                If vTag.class = "btn-edit" And vTag.target = "_blank" And vTag.title = "修改" Then List2.AddItem vTag.href
'            End If
        Case "B"
            'A SPAN　B  B B
            
'            If UCase(vTag_2.TagName) = "A" And _
'               UCase(vTag_1.TagName) = "SPAN" And _
'               UCase(vTag1.TagName) = "B" And _
'               UCase(vTag2.TagName) = "B" Then
'                Me.Caption = "当前的任务有" & vTag.innerhtml & "个!"
'            End If

            'http://192.168.0.8:83/app1/TaskLadingCn/List.aspx?k=&RearchType=0&UId=0&KfId=4986&MgId=0&mgbm=0&bumen=0&followup=&FState=0&tdtype=-1&FSpeed=0&FMgSpeed=0&FKfSpeed=1&attr=0&AttrBus=0&selDate=0&strDate=&endDate=
            'TD A  IMG  TD P
        Case "IMG"
            If UCase(vTag_2.TagName) = "TD" And _
               UCase(vTag_1.TagName) = "A" And _
               UCase(vTag1.TagName) = "TD" And _
               UCase(vTag2.TagName) = "p" Then
                'List2.AddItem vTag_1.innerhtml
                Debug.Print vTag_1.innerhtml
                Debug.Print vTag.src
            End If
        Case "EM"
            If UCase(vTag_2.TagName) = "A" And _
               UCase(vTag_1.TagName) = "LI" And _
               UCase(vTag1.TagName) = "INPUT" And _
               UCase(vTag2.TagName) = "LI" Then
                pages = vTag.innerhtml
            End If
        End Select
    Next
End Sub

Private Sub web_NewWindow2(index As Integer, ppDisp As Object, Cancel As Boolean)
    Dim i
    For i = 1 To web.UBound
        If web(i).Tag Then Set ppDisp = web(i).Object: showweb (i): pages = "已加载...": Exit Sub
    Next
    pages = "未加载..."
    Cancel = True
End Sub

Private Sub showweb(index As Long)
    Dim i As Long
    For i = 1 To web.UBound
        web(i).Visible = False
        List2.Visible = False
    Next
    web(index).Visible = True
End Sub
