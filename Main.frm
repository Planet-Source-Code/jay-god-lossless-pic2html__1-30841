VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Main 
   Caption         =   "Lossless Pic to HTML Conversion"
   ClientHeight    =   4830
   ClientLeft      =   2250
   ClientTop       =   2520
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   322
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   Begin VB.TextBox txtPath2 
      Height          =   285
      Left            =   3855
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.PictureBox picPercent 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080C0FF&
      DrawMode        =   6  'Mask Pen Not
      Height          =   345
      Left            =   60
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   488
      TabIndex        =   11
      Top             =   630
      Width           =   7380
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Load HTML Code"
      Height          =   375
      Left            =   5625
      TabIndex        =   10
      Top             =   3435
      Width           =   1785
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Center Image"
      Height          =   195
      Left            =   5970
      TabIndex        =   9
      Top             =   315
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Save HTML Code"
      Height          =   375
      Left            =   90
      TabIndex        =   7
      Top             =   3435
      Width           =   1785
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Copy HTML Code"
      Height          =   375
      Left            =   1905
      TabIndex        =   6
      Top             =   3435
      Width           =   1785
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   360
      Left            =   3510
      TabIndex        =   5
      Top             =   240
      Width           =   360
   End
   Begin RichTextLib.RichTextBox txtHTMLSource 
      Height          =   2400
      Left            =   60
      TabIndex        =   4
      Top             =   990
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   4233
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Main.frx":0000
   End
   Begin VB.PictureBox picPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   90
      Picture         =   "Main.frx":0082
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   42
      TabIndex        =   3
      Top             =   4125
      Width           =   630
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Generate Code!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3930
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   285
      Width           =   3360
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "www.necrocosm.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   30
      TabIndex        =   8
      Top             =   3825
      Width           =   7425
   End
   Begin VB.Image imgPic 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   5730
      Picture         =   "Main.frx":0CC6
      Top             =   30
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label1 
      Caption         =   "Choose a picture file to convert to HTML:"
      Height          =   195
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   3300
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Common Dialog stuff:
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (lpopenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFiletitle As String
    nMaxfileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

'for shelling:
Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

'for getting pixel color:
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long

'For timeout:
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub Timeout(dblSeconds As Double)
Dim hCurrent As Long
    
dblSeconds = dblSeconds * 1000
hCurrent = GetTickCount()

Do While GetTickCount() - hCurrent < dblSeconds
    DoEvents
Loop
End Sub

Public Sub SetPercent(Picture As Control, ByVal Percent)
Dim num As String

Picture.Cls
Picture.ScaleHeight = 100
Picture.ScaleWidth = 100
num = Format$(Percent, "###") + "%"
num = Val(num) & "%"

'    Picture.ScaleWidth = 100
num$ = Format$(Percent, "###") + "%"
 Picture.CurrentX = 50 - Picture.TextWidth(num) / 2
 Picture.CurrentY = (Picture.ScaleHeight - Picture.TextHeight(num)) / 2
 Picture.Print num
Picture.Line (0, 0)-(Percent, Picture.ScaleHeight), , BF

Picture.Refresh
Timeout 0.001 'pause to display it
End Sub

Private Function GetHex(intLongValue As Long) As String
Dim intBlue As Long
Dim intGreen As Long
Dim intRed As Long
Dim strBlue As String 'hex
Dim strRed As String 'hex
Dim strGreen As String 'hex

    If intLongValue >= 65536 Then
        intBlue = Int(intLongValue / 65536)
        intLongValue = intLongValue - (65536 * intBlue)
    End If

    If intLongValue >= 256 Then
        intGreen = Int(intLongValue / 256)
        intLongValue = intLongValue - (256 * intGreen)
    End If
    
    intRed = intLongValue

    strBlue = Hex(intBlue)
    strRed = Hex(intRed)
    strGreen = Hex(intGreen)
    
    If Len(strBlue) < 2 Then strBlue = "0" & strBlue
    If Len(strRed) < 2 Then strRed = "0" & strRed
    If Len(strGreen) < 2 Then strGreen = "0" & strGreen

    GetHex = strBlue & strRed & strGreen

End Function


Public Function LaunchSite(strUrl As String) As Long

    Dim lhWnd As Long
    Dim lAns As Long
    ' Execute the url
    lAns = ShellExecute(lhWnd, "open", strUrl, vbNullString, vbNullString, 3) '3=MAXIMIZE WINDOW
   
    OpenLocation = lAns ' return returnval

End Function


Private Sub SaveAndLoadHTML()
'This will save the textbox to a file and launch it in the browser
Dim strText As String

If Check1.Value = 1 Then 'center it
    txtHTMLSource.Text = "<HTML><HEAD><TITLE>http://www.necrocosm.com</TITLE></HEAD><BODY BGCOLOR=""#FFFFFF""><CENTER>" & txtHTMLSource.Text
    txtHTMLSource.Text = txtHTMLSource.Text & "</CENTER></BODY></HTML>"
Else
    txtHTMLSource.Text = "<HTML><HEAD><TITLE>http://www.necrocosm.com</TITLE></HEAD><BODY BGCOLOR=""#FFFFFF"">" & txtHTMLSource.Text
    txtHTMLSource.Text = txtHTMLSource.Text & "</BODY></HTML>"
End If


Open "C:\TestHTML.html" For Output As #11
Close #11

Open "C:\TestHTML.html" For Binary As #130
    strText = txtHTMLSource.Text
    Put #130, 1, strText
Close #130

Call LaunchSite("C:\TestHTML.html")
End Sub

Private Sub Command1_Click()
Dim strHTMLCode As String
Dim strHexColor As String 'Hex code for the current pixel
Dim intPicWidth As Integer
Dim intPicHeight As Integer
Dim intColor As Long
Dim intLastColor As Long
Dim intNextColor As Long
Dim picHDC As Long
Dim intRowNumber As Integer
Dim i As Integer
Dim j As Integer
Dim x As Integer

Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False

picHDC = picPic.hDC
intPicWidth = picPic.Width
intPicHeight = picPic.Height
intRowNumber = 1

'BEGIN Algorithm...
strHTMLCode = "<TABLE WIDTH=" & intPicWidth & " HEIGHT=" & intPicHeight & " CELLSPACING=0 CELLPADDING=0 BORDER=0>"

For i = 0 To intPicHeight - 1 'vertical rows
    DoEvents
    strHTMLCode = strHTMLCode & "<TR HEIGHT=1>" '1 pixel height
    For j = 0 To intPicWidth - 1 'horizontal rows
        DoEvents
        intColor = GetPixel(picHDC, j, i)
        If j < intPicWidth - 1 Then 'Not the last pixel
            intNextColor = GetPixel(picHDC, j + 1, i)
            If intColor = intNextColor Then 'Check total amount of repeating colors
                intRowNumber = 2 'there's definitely 2, check for more:
                For x = j + 2 To intPicWidth - 1 'check remaining row
                    intNextColor = GetPixel(picHDC, x, i)
 '                   MsgBox x & "," & i & "=" & intNextColor, 64, "X LOOP"
                    If intColor = intNextColor Then 'same again
                        If x <> intPicWidth - 1 Then
                            intRowNumber = intRowNumber + 1
                        Else
                            intRowNumber = intRowNumber + 1
                            j = intPicWidth - 1
                        End If
                    Else 'not the same
                        j = x - 1
                        Exit For
                    End If
                Next 'x
            End If
        End If

    'MsgBox intRowNumber, 64, j & ",i:" & i
        If intRowNumber = 1 Then
            strHexColor = Hex(intColor)
            strHTMLCode = strHTMLCode & "<TD WIDTH=" & intRowNumber & " BGCOLOR=""" & strHexColor & """></TD>"
        Else
            strHexColor = Hex(intColor)
            strHTMLCode = strHTMLCode & "<TD WIDTH=" & intRowNumber & " BGCOLOR=""" & strHexColor & """ COLSPAN=" & intRowNumber & "></TD>"
            intRowNumber = 1
        End If
    Next 'j
    
    strHTMLCode = strHTMLCode & "</TR>"
    Call SetPercent(picPercent, (i / (intPicHeight - 1)) * 100)
Next 'i

strHTMLCode = strHTMLCode & "</TABLE>"
'END Algorithm.

txtHTMLSource.Text = strHTMLCode

Call SaveAndLoadHTML

Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
End Sub

Private Sub Command2_Click()
On Error Resume Next

Clipboard.SetText txtHTMLSource.Text

MsgBox "The HTML Source has been copied to the clipboard.", 64, "Copied!"
End Sub

Private Sub Command3_Click()
Dim i As Byte
Dim j As Byte
Dim intFreeFile As Byte
Dim strText As String 'The entire file of params
Dim strParam As String 'Current Parameter
'...+
Dim rc As Long 'Return Data Holder
Dim cdFindFile As OPENFILENAME
Dim FileOpen As String
Const MAX_BUFFER_LENGTH = 256
On Error GoTo ErrHandler

'Set dialog's stuff:
cdFindFile.hwndOwner = hwnd
cdFindFile.hInstance = App.hInstance
cdFindFile.lpstrTitle = "Open Image..."
cdFindFile.lpstrInitialDir = CurDir() 'App.Path
cdFindFile.lpstrFilter = "Image Files (.bmp, .gif, .jpg)" & Chr$(0) & "*.BMP;*.JPG;*.GIF" & Chr$(0)
cdFindFile.nFilterIndex = 1
cdFindFile.flags = &H4 + &H800 'Don't show read only + Path must exist
cdFindFile.lpstrFile = String(MAX_BUFFER_LENGTH, Chr$(0))
cdFindFile.nMaxFile = MAX_BUFFER_LENGTH - 1
cdFindFile.lpstrFiletitle = cdFindFile.lpstrFile
cdFindFile.nMaxfileTitle = MAX_BUFFER_LENGTH - 1
cdFindFile.lStructSize = Len(cdFindFile)

rc = GetOpenFileName(cdFindFile)

If rc Then
    FileOpen = Left$(cdFindFile.lpstrFile, cdFindFile.nMaxFile)
    txtPath.Text = FileOpen
    FileOpen = txtPath.Text 'Remove nulls

    strSP3Path = FileOpen 'SET THE PROJECT PATH
Else 'Cancel was pressed
    FileOpen = ""
    Exit Sub
End If

imgPic.Picture = LoadPicture(FileOpen)
picPic.Width = imgPic.Width
picPic.Height = imgPic.Height
picPic.Picture = imgPic.Picture


Exit Sub
ErrHandler:
MsgBox "An Image File cannot be opened due to the following error:" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & " " & Err.Number & " - " & Err.Description, 16, "Error!"
End Sub

Private Sub Command4_Click()
Dim i As Byte
Dim j As Byte
Dim intFreeFile As Byte
Dim strText As String 'The entire file of params
Dim strParam As String 'Current Parameter
'...+
Dim rc As Long 'Return Data Holder
Dim cdFindFile As OPENFILENAME
Dim FileOpen As String
Const MAX_BUFFER_LENGTH = 256
On Error GoTo ErrHandler

'Set dialog's stuff:
cdFindFile.hwndOwner = hwnd
cdFindFile.hInstance = App.hInstance
cdFindFile.lpstrTitle = "Save HTML Source..."
cdFindFile.lpstrInitialDir = CurDir() 'App.Path
cdFindFile.lpstrFilter = "HTML Source Files (*.htm, *.html)" & Chr$(0) & "*.htm;*.html" & Chr$(0)
cdFindFile.nFilterIndex = 1
cdFindFile.flags = &H4 + &H80000 + &H200000 + &H2 'Hide Read Only ;explorer ; OFN_LONGNAMES long names + overwrite prompt
cdFindFile.lpstrFile = String(MAX_BUFFER_LENGTH, Chr$(0))
cdFindFile.nMaxFile = MAX_BUFFER_LENGTH - 1
cdFindFile.lpstrFiletitle = cdFindFile.lpstrFile
cdFindFile.nMaxfileTitle = MAX_BUFFER_LENGTH - 1
cdFindFile.lStructSize = Len(cdFindFile)

rc = GetSaveFileName(cdFindFile)

If rc Then
    FileOpen = Left$(cdFindFile.lpstrFile, cdFindFile.nMaxFile)
    txtPath2.Text = FileOpen
    FileOpen = txtPath2.Text 'Remove nulls

Else 'Cancel was pressed
    FileOpen = ""
    Exit Sub
End If

If Right$(FileOpen, 4) <> ".htm" Or Right$(FileOpen, 5) <> ".html" Then
    FileOpen = FileOpen & ".html"
End If

Open FileOpen For Output As #11
Close #11

Open FileOpen For Binary As #130
    strText = txtHTMLSource.Text
    Put #130, 1, strText
Close #130

MsgBox "Save successfully as:" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & FileOpen, 64, "Saved!"

Exit Sub
ErrHandler:
MsgBox "An HTML File cannot be saved due to the following error:" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & " " & Err.Number & " - " & Err.Description, 16, "Error!"
End Sub

Private Sub Command5_Click()
Dim strText As String
On Error Resume Next

Close

Open "C:\TestHTML.html" For Output As #11
Close #11

Open "C:\TestHTML.html" For Binary As #130
    strText = txtHTMLSource.Text
    Put #130, 1, strText
Close #130

Call LaunchSite("C:\TestHTML.html")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Open "C:\TestHTML.html" For Output As #13
Close #13

End
End Sub

Private Sub Label2_Click()
Call LaunchSite("http://www.necrocosm.com")
End Sub
