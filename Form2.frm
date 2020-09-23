VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form2 
   Caption         =   "Super Sweet Donkey Stick"
   ClientHeight    =   1215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6375
   LinkTopic       =   "Form2"
   ScaleHeight     =   1215
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Make Sextastic HTML Picture From .bmp Picture!"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   6135
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5040
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   5880
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   ".bmp path:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Sub Command1_Click()
cd1.Filter = ".bmp|*.bmp"
cd1.ShowOpen
If cd1.FileName = "" Then Exit Sub
Text1.Text = cd1.FileName
Form1.pb1.Picture = LoadPicture(cd1.FileName)
End Sub

Private Sub Command2_Click()
Form2.Caption = "Super Sweet Donkey Stick - Creating File..."
Dim aPix() As String, lPX As Long, bR As Byte, bG As Byte, bB As Byte
Dim lWidth As Long, oldPix As String, lCur As Long
' reset the file and add the basics
Open App.Path & "/MyPic.html" For Output As #1
Print #1, "<html>"
Print #1, "<body bgcolor=000000>"
Print #1, "<table border=0 cellpadding=0 cellspacing=0 bgcolor=ffffff>"
Print #1, "<tr height=1>"
ReDim aPix(0 To Form1.pb1.ScaleWidth, 0 To Form1.pb1.ScaleHeight)
lHeight = 1
lWidth = 1
' setup oldpix here outside the loop to stop from having 1 more command every time, faster in the long run
lPX = GetPixel(Form1.pb1.hdc, 0, 0)
bR = lPX And 255
bG = (lPX \ 256) And 255
bB = (lPX \ 65536) And 255
oldPix = rgbtohex(bR, bG, bB)
For y = 0 To Form1.pb1.ScaleHeight - 1
    For x = 0 To Form1.pb1.ScaleWidth - 1
        lPX = GetPixel(Form1.pb1.hdc, x, y)
        bR = lPX And 255
        bG = (lPX \ 256) And 255
        bB = (lPX \ 65536) And 255
        aPix(x, y) = rgbtohex(bR, bG, bB)
        If oldPix = aPix(x, y) Then
            lWidth = lWidth + 1
            Else
            'paint a line, the width before the pixel color was changed
            'that way it makes lines not just pixel by pixel unless the line
            'is only 1 pixel long, the height always goes by 1 pixel tho
            Print #1, "<td width=" & lWidth & " bgcolor=" & oldPix & "></td>"
            lWidth = 1
        End If
        oldPix = aPix(x, y)
        lCur = lCur + 1
        Call UpdateProgress(lCur, ((Form1.pb1.ScaleHeight - 1) * (Form1.pb1.ScaleWidth - 1)))
    Next
    Print #1, "<td width=" & lWidth & " bgcolor=" & oldPix & "></td>"
    lWidth = 1
    ' print the next height to start at
    If y <> Form1.pb1.ScaleHeight - 1 Then
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 bgcolor=ffffff>"
        Print #1, "<tr height=1>"
    End If
Next
Print #1, "</table>"
Print #1, "</body>"
Print #1, "</html>"
Close #1
Form2.Caption = "Super Sweet Donkey Stick - Done!"
End Sub

Sub UpdateProgress(lProg As Long, lMax As Long)
'for all you retards out there, the way to find percentage
'
'lprog       x
'-----     -----
' max       100
'
'or
'lprog * 100
'then
'answer / max

Dim lTmp As Long, sTmp As String
lTmp = ((lProg * 100) / lMax)
'Form2.Caption = "Super Sweet Donkey Stick - Creating HTML File: " & tmp & "%"
For x = 0 To lTmp
    sTmp = sTmp & "||"
Next
Label2.Caption = sTmp
DoEvents
End Sub

'not my function, got it offa pscode imlazy
'http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=49849&lngWId=1
Public Function rgbtohex(r As Byte, g As Byte, b As Byte) As String

'input format = 255,255,255

'get the r value
If r < 16 Then
hex1 = 0 & Hex(r)
Else
hex1 = Hex(r)
End If


'get the g value
If r < 16 Then
hex2 = 0 & Hex(g)
Else
hex2 = Hex(g)
End If


'get the b value
If b < 16 Then
hex3 = 0 & Hex(b)
Else
hex3 = Hex(b)
End If

rgbtohex = "#" & hex1 & hex2 & hex3
End Function


Private Sub Form_Load()
Form1.Show
End Sub
