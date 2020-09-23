Attribute VB_Name = "Module1"
Global Top1 As Integer
Global Left1 As Integer
Global X1 As Integer
Global Y1 As Integer
Public Sub RealLabel()
   Form2.Label2.AutoSize = True
   Form2.Width = Form2.Label2.Width
   Form2.Height = Form2.Label2.Height
End Sub
Public Sub Tooltip()
  Form2.Top = Y1 + Top1 + Form1.Top + 300
  Form2.Left = X1 + Left1 + Form1.Left - 300
  Form2.Show 0
  Form2.Label2.Caption = ""
End Sub

