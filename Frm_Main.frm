VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Frm_Main 
   Caption         =   "������ͼ"
   ClientHeight    =   9390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19035
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   19035
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton lg 
      Caption         =   "login"
      Height          =   270
      Left            =   12660
      TabIndex        =   17
      Top             =   105
      Width           =   735
   End
   Begin VB.Timer getcb 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8910
      Top             =   4455
   End
   Begin SHDocVwCtl.WebBrowser mm 
      Height          =   720
      Left            =   10920
      TabIndex        =   16
      Top             =   1335
      Visible         =   0   'False
      Width           =   4530
      ExtentX         =   7990
      ExtentY         =   1270
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
   Begin VB.Timer getass 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8910
      Top             =   4455
   End
   Begin ������Ʒ��ͼ.Frm_Tols Frm 
      Height          =   240
      Left            =   9390
      Top             =   4575
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin VB.CommandButton loginassist 
      Caption         =   "�̻�����"
      Height          =   300
      Left            =   13455
      TabIndex        =   15
      Top             =   90
      Width           =   930
   End
   Begin VB.CommandButton pic 
      Caption         =   "ͼƬ"
      Height          =   300
      Left            =   10905
      TabIndex        =   14
      Top             =   90
      Width           =   600
   End
   Begin VB.CommandButton manager 
      Caption         =   "��Ʒ"
      Height          =   300
      Left            =   10170
      TabIndex        =   13
      Top             =   90
      Width           =   600
   End
   Begin VB.CommandButton alibaba 
      Caption         =   "1688"
      Height          =   300
      Left            =   9435
      TabIndex        =   12
      Top             =   90
      Width           =   600
   End
   Begin VB.CommandButton oa 
      Caption         =   "OA"
      Height          =   300
      Left            =   8700
      TabIndex        =   11
      Top             =   90
      Width           =   600
   End
   Begin VB.CommandButton lookitem 
      Caption         =   "�鿴��Ʒ"
      Height          =   300
      Left            =   11640
      TabIndex        =   10
      Top             =   90
      Width           =   945
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   915
      Index           =   2
      Left            =   9750
      TabIndex        =   9
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
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   915
      Index           =   1
      Left            =   8580
      TabIndex        =   2
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
   Begin VB.ListBox List2 
      Height          =   780
      Left            =   45
      TabIndex        =   7
      Top             =   9315
      Width           =   17145
   End
   Begin VB.TextBox urlT 
      Height          =   270
      Index           =   2
      Left            =   8700
      TabIndex        =   6
      Text            =   "2"
      Top             =   480
      Width           =   7710
   End
   Begin VB.TextBox urlT 
      Height          =   270
      Index           =   1
      Left            =   945
      TabIndex        =   4
      Text            =   "1"
      Top             =   480
      Width           =   7710
   End
   Begin VB.TextBox urlT 
      Height          =   270
      Index           =   0
      Left            =   945
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   105
      Width           =   7710
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   7665
      Index           =   0
      Left            =   45
      TabIndex        =   1
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
   Begin VB.Label pages 
      AutoSize        =   -1  'True
      Caption         =   "ҳ��"
      Height          =   180
      Left            =   14550
      TabIndex        =   8
      Top             =   150
      Width           =   360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "��ҳ����:"
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   525
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��ҳ����:"
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
Private Savetime As Double
Dim strUrl As String
Dim uid As String
Dim username As String
Dim password As String
Dim showmsg As Boolean
Dim gsmc As String

Dim assistpid As Long
Dim assisthWnd As Long

Private Sub alibaba_Click()
    web(0).Navigate2 "http://work.1688.com/home/page/index.htm#nav/home"
End Sub

'
Private Sub Form_Load()
    web(0).Navigate "http://192.168.0.8:83/"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim lefthg
    web(0).Top = 900
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
End Sub

Private Sub getfp(webb As WebBrowser)
    On Error Resume Next
    Dim i, J, vDoc
    Dim ix As Long
    Dim itemname, itemurl
    ix = webb.index
    Set vDoc = webb.Document
    itemname = resetfilename(vDoc.getelementsbytagname("input")("subject").Value)
    ERR.clear
    itemurl = vDoc.getelementsbytagname("input")("pictureUrl").Value
    If ERR <> 0 Then
        itemurl = vDoc.getelementsbytagname("input")("pictureUrl")(0).Value
    End If

    If InStr(1, itemurl, "http") <> 0 And InStr(1, itemurl, "jpg") <> 0 And InStr(1, itemurl, ".com//") = 0 Then
        If Frm_Download.UName.AddItemNotSame(itemurl) Then
            If Not (Frm_Download.SName.AddItemNotSame(resetfilename(Trim(itemname) & ".jpg"))) Then
                Frm_Download.SName.AddItemNotSame resetfilename((Trim(itemname) & i) & ".jpg")
            End If
        End If
    End If

End Sub

Private Function resetfilename(ByVal name As String) As String
    On Error Resume Next
    name = clear(name, "/")
    name = clear(name, "\")
    name = clear(name, "*")
    name = clear(name, "?")
    name = clear(name, "<")
    name = clear(name, ">")
    name = clear(name, ":")
    resetfilename = name
End Function

Private Function clear(name As String, P As String) As String
    clear = Replace(name, P, "")
End Function

Private Sub Form_Unload(Cancel As Integer)
    Unload Frm_Download
    End
End Sub

Private Sub getass_Timer()
    On Error Resume Next
    Dim Cname As String
    assisthWnd = Frm.GethWndByClass("#32770")
    '"����Ͱ��̻������¼"
    If assisthWnd <> 0 Then
        Debug.Print hWnd
        Cname = Frm.GetTitle(assisthWnd)
        Debug.Print Cname
        Frm.SetFedWnd assisthWnd
        assistpid = Frm.GetPidByhWnd(assisthWnd)
        getass.Enabled = False
        getcb.Enabled = True
    End If
End Sub

Private Sub getcb_Timer()
    On Error Resume Next
    Dim hWnd As Long
    Dim mcname As String
    hWnd = Frm.GetMouseWindowhWnd
    If Frm.GetPidByhWnd(hWnd) = assistpid Then
        mcname = Frm.GetClassName(hWnd)
        If mcname = "ComboBox" Then
            Frm.SetFedWnd hWnd
            Debug.Print "�ѻ�ȡ�����" & hWnd
            Frm.hWndKeyPress hWnd, vbKeyLButton, System
            Frm.hWndKeyPress assisthWnd, vbKeyTab, System
            SendKeys "123"
            Frm.hWndKeyPress assisthWnd, vbKeyTab, System
            SendKeys "123"
            getcb.Enabled = False
        End If
    End If
End Sub

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
                If vTag.Class = "next" Then vTag.Click
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

Private Sub lg_Click()
    getass.Enabled = True
End Sub

Private Sub loginassist_Click()
    On Error Resume Next
    Dim jj
    'If InStr(1, mm.LocationURL, "alilogin.aspx") Then
    If uid <> "" And username <> "" Then
        mm.Navigate "http://192.168.0.8:83/"
        pages.Caption = "��ȡ������..."
        jj = MsgBox("��˾����: " & gsmc & vbCrLf & vbCrLf & "��ǰ�û�UID: " & uid & vbCrLf & "��ǰ�û�����: " & username & vbCrLf _
                  & "��ǰ�û�����: ********" & vbCrLf & vbCrLf & "�Ƿ������¼?", vbExclamation + vbYesNo)
        If jj = vbYes Then
            Do Until Not mm.Busy
                Sleep 10
            Loop
            mm.Navigate "javascript:document.getElementById('sixiAX').GetPW('" + uid + "','192.168.0.8:83')"
            Dim password
            showmsg = True
        End If
    Else
        MsgBox "���ȴ�һ�ΰ�������!"
    End If
End Sub

Private Sub lookitem_Click()
    Frm_Download.Show
End Sub

Private Sub manager_Click()
    web(0).Navigate2 "http://offer.1688.com/offer/manage.htm?show_type=valid&tracelog=work_1_m_orderManage"
End Sub

Private Sub mm_DocumentComplete(ByVal pDisp As Object, url As Variant)
    On Error Resume Next
    If showmsg Then
        password = mm.Document.body.innerhtml
        Clipboard.clear
        Clipboard.SetText password
        MsgBox "�����ȡ�ɹ�,����1���Ӵ��̻�����,�����Զ���¼!", vbInformation, "�̻������¼"
        getass.Enabled = True
    End If
    showmsg = False
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
    'On Error Resume Next
    Dim target, Title, Class
    Dim itemurl As String
    Dim itemname As String
    web(index).Silent = True
    web(index).Tag = True
    urlT(index).Enabled = True
    urlT(index).ForeColor = vbBlue
    Me.Caption = "Load Complete"
    showweb (index)
    Dim vDoc, vTag_2, vTag_1, vTag, vTag1, vTag2, vTXT
    Dim i As Integer
    Set vDoc = web(index).Document
    'On Error Resume Next
    If InStr(1, web(index).LocationURL, "alilogin.aspx") Then
        uid = Mid(web(index).LocationURL, InStr(1, web(index).LocationURL, "?id=") + 4)
        username = vDoc.getelementsbytagname("input")("TPL_username").Value
    End If
    For i = 2 To vDoc.All.length - 1
        On Error Resume Next
        Set vTag_2 = vDoc.All(i - 2)
        Set vTag_1 = vDoc.All(i - 1)
        Set vTag = vDoc.All(i)
        Set vTag1 = vDoc.All(i + 1)
        Set vTag2 = vDoc.All(i + 2)
        Select Case UCase(vDoc.All(i).TagName)
        Case "TD"
        Case "A"
            '��Ʒ�б�������ȡ��Ϣ
            If vTag.Title = "�鿴��˾��ϸ��Ϣ" Then gsmc = vTag1.innerhtml
            Dim st As Long
            Dim en As Long
            If UCase(vTag_2.TagName) = "INPUT" And _
               UCase(vTag_1.TagName) = "TD" And _
               UCase(vTag1.TagName) = "IMG" And _
               UCase(vTag2.TagName) = "TD" Then
                st = InStr(1, vTag.innerhtml, "data-lazyload-src=""") + Len("data-lazyload-src=""")    'data-lazyload-src="http://
                st = InStr(st + 1, vTag.innerhtml, "/") + Len("/")
                en = InStr(st, vTag.innerhtml, "jpg") + 3
                itemurl = Mid(vTag.innerhtml, st, en - st)
                itemurl = urlreset(itemurl)
                Debug.Print itemurl
                If InStr(1, itemurl, "http") <> 0 And InStr(1, itemurl, "jpg") <> 0 And InStr(1, itemurl, ".com//") = 0 Then
                    If Frm_Download.UName.AddItemNotSame(itemurl) Then
                        If Not (Frm_Download.SName.AddItemNotSame(resetfilename(Trim(vTag.Title) & ".jpg"))) Then
                            Frm_Download.SName.AddItemNotSame resetfilename((Trim(vTag.Title) & i) & ".jpg")
                        End If
                    End If
                End If
            End If
        Case "META"
            '��Ʒչʾ����ֱ�ӻ�ȡ��ͼ��Ϣ
            If vTag.Property = "og:image" And vTag1.Property = "og:title" Then
                itemurl = urlreset(vTag.content)
                itemname = vTag1.content
                If InStr(1, itemurl, "http") <> 0 And InStr(1, itemurl, "jpg") <> 0 And InStr(1, itemurl, ".com//") = 0 Then
                    If Frm_Download.UName.AddItemNotSame(itemurl) Then
                        If Not (Frm_Download.SName.AddItemNotSame(resetfilename(Trim(itemname) & ".jpg"))) Then
                            Frm_Download.SName.AddItemNotSame resetfilename((Trim(itemname) & i) & ".jpg")
                        End If
                    End If
                End If
            End If
        End Select
    Next
End Sub

Private Sub web_NewWindow2(index As Integer, ppDisp As Object, Cancel As Boolean)
    On Error Resume Next
    Dim i
    For i = 1 To web.UBound
        If web(i).Tag Then Set ppDisp = web(i).object: showweb (i): pages = "�Ѽ���...": Exit Sub
    Next
    pages = "δ����..."
    Cancel = True
End Sub

Public Function urlreset(ByVal url As String) As String
    Dim st, en
    'Debug.Print url
    st = InStr(1, url, "http://") + Len("http://")
    st = InStr(st + 1, url, "/") + Len("/")
    en = InStr(st, url, "jpg") + 3
    url = Mid(url, st, en - st)
    url = Replace(url, ".310x310", "")
    url = Replace(url, ".64x64", "")
    url = "http://i01.c.aliimg.com/" & url
    urlreset = url
    'Debug.Print url
End Function

Private Sub showweb(index As Long)
    Dim i As Long
    For i = 1 To web.UBound
        web(i).Visible = False
        List2.Visible = False
    Next
    web(index).Visible = True
End Sub
