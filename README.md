#LoadFirstPic
代码1：

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
Dim frm As Form1
Set frm = New Form1
frm.Visible = True
Set ppDisp = frm.WebBrowser1.object
End Sub

代码2：

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
Cancel = True
WebBrowser1.Navigate2 WebBrowser1.Document.activeElement.href
End Sub

代码3：

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
On Error Resume Next
Dim frmWB As Form1
Set frmWB = New Form1
frmWB.WebBrowser1.RegisterAsBrowser = True
Set ppDisp = frmWB.WebBrowser1.object
frmWB.Visible = True
frmWB.Top = Form1.Top
frmWB.Left = Form1.Left
frmWB.Width = Form1.Width
frmWB.Height = Form1.Height
End Sub

代码4：这个最好用了

Dim WithEvents Web_V1 As SHDocVwCtl.WebBrowser_V1

PrivateSub Form_Load()
    Set Web_V1 = WebBrowser1.Object
End Sub
    
PrivateSub Web_V1_NewWindow(ByVal URL AsString, ByVal Flags AsLong, ByVal TargetFrameName AsString, PostData As Variant, ByVal Headers AsString, Processed AsBoolean)
    Processed =True
    WebBrowser1.Navigate URL
End Sub