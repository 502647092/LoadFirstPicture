VERSION 5.00
Begin VB.Form Frm_Download 
   Caption         =   "��ͼ����"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   11565
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton clear 
      Caption         =   "���"
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   285
      Width           =   1020
   End
   Begin ������Ʒ��ͼ.TzProgressBar pb 
      Height          =   255
      Left            =   1080
      Top             =   330
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "�ܽ���"
      BackColor       =   8438015
      StartColor      =   8438015
   End
   Begin ������Ʒ��ͼ.TzDownload dl 
      Height          =   250
      Left            =   1065
      Top             =   60
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   450
      ForeColor       =   16777088
   End
   Begin VB.CommandButton dlc 
      Caption         =   "����"
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1020
   End
   Begin ������Ʒ��ͼ.TzListBox UName 
      Height          =   1170
      Left            =   0
      TabIndex        =   0
      Top             =   2055
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   2064
   End
   Begin ������Ʒ��ͼ.TzListBox SName 
      Height          =   1335
      Left            =   -15
      TabIndex        =   1
      Top             =   720
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   2355
   End
End
Attribute VB_Name = "Frm_Download"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'=================================Sleep========================================
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Savetime As Double
Private sd As Boolean

Private Sub clear_Click()
    SName.clear
    UName.clear
End Sub

Private Sub dl_OnFinished(ByVal Result As Boolean)
    sd = Result
End Sub

Private Sub dlc_Click()
    Dim i
    Dim folder As String
    Dim usetime As Double
    If folder = "" Then folder = InputBox("������ ����-��ͼ-��˾����-�����˺�-�ᵥ������!", , Format(Now, "m.d") & "-��ͼ-��˾����-�����˺�-�ᵥ������")
    If folder = "" Then folder = Format(Now, "m.d") & "-��ͼ-��˾����-�����˺�-�ᵥ������"
    usetime = timeGetTime
    For i = 0 To UName.ListCount - 1
red:
        pb.Change i, "������ ����:  " & i & "/" & pb.BarMax
        UName.ListIndex = i
        SName.ListIndex = i
        dl.FileDownload UName.List(i), App.Path & "\" & folder & "\" & Trim(SName.List(i))    ' & ".jpg"
        Do
            Sleep 50
        Loop Until dl.IsFree
        If Not sd Then GoTo red
        'Debug.Print Replace(Trim(SName.List(i)), " ", "")
    Next
    usetime = Format((timeGetTime - usetime) / 1000, "0.00")
    pb.Change pb.BarMax, "������� ������" & pb.BarMax & "����Ʒ��ͼ ��ʱ" & usetime & "��!", &H80FF80
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Me.Hide
    Cancel = True
End Sub

'itempicurl (ix), App.Path & "\" & folder.Text & "\" & itemname(ix).Text & ".jpg"
Private Sub Form_Resize()
    On Error Resume Next
    
    SName.Left = 5
    SName.Top = 600
    SName.Height = Me.Height - 600
    SName.Width = Me.Width / 2 - 10
    
    UName.Left = Me.Width / 2 + 10
    UName.Top = 600
    UName.Height = Me.Height - 600
    UName.Width = Me.Width / 2 - 10
    
    dl.Left = dlc.Left + dlc.Width + 10
    dl.Width = Me.Width - dl.Left - 20
    dl.Top = 25
    
    pb.Left = dlc.Left + dlc.Width + 10
    pb.Width = Me.Width - dl.Left - 20
    pb.Top = dl.Top + dl.Height + 50
End Sub

Public Sub Sleep(n As Long)
    Savetime = timeGetTime
    While timeGetTime < Savetime + n
        DoEvents
    Wend
End Sub

Private Sub SName_dblClick()
    InputBox "����ѡ��Ĳ�Ʒ��������:", , SName.List(SName.ListIndex)
End Sub

Private Sub UName_AddItem()
    pb.BarMax = UName.ListCount
    pb.Change pb.BarMax, "��ɨ�赽��Ʒ��Ϣ" & UName.ListCount & "��"
End Sub

Private Sub UName_dblClick()
    InputBox "����ѡ��Ĳ�Ʒ��������:", , UName.List(UName.ListIndex)
End Sub
