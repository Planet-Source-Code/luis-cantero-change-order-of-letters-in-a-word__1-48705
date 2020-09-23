VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WordChanger"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   495
      Left            =   10080
      TabIndex        =   2
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox txtTo 
      Height          =   2655
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3120
      Width           =   11055
   End
   Begin VB.TextBox txtFrom 
      Height          =   2655
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmMain.frx":0000
      Top             =   120
      Width           =   11055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdConvert_Click()

  Dim strWord As String
  Dim strTemp As String
  Dim i As Integer
  Dim intSep1 As Integer
  Dim intSep2 As Integer
  Dim intSep3 As Integer
  Dim intSep4 As Integer
  Dim intSep5 As Integer

    If Len(txtFrom) < 4 Then Exit Sub

    cmdConvert.Enabled = False

    strTemp = txtFrom

    For i = 1 To Len(txtFrom)

        'Search for delimiter
        intSep1 = InStr(i + 1, txtFrom, " ")
        intSep2 = InStr(i + 1, txtFrom, ".")
        intSep3 = InStr(i + 1, txtFrom, ",")
        intSep4 = InStr(i + 1, txtFrom, ";")
        intSep5 = InStr(i + 1, txtFrom, vbCrLf)

        'Correct if delimiter was not found
        If intSep1 = 0 Then intSep1 = intSep2
        If intSep1 = 0 Then intSep1 = intSep3
        If intSep1 = 0 Then intSep1 = intSep4
        If intSep1 = 0 Then intSep1 = intSep5

        'Correct delimiter position
        If intSep2 < intSep1 And intSep2 > 0 Then intSep1 = intSep2
        If intSep3 < intSep1 And intSep3 > 0 Then intSep1 = intSep3
        If intSep4 < intSep1 And intSep4 > 0 Then intSep1 = intSep4
        If intSep5 < intSep1 And intSep5 > 0 Then intSep1 = intSep5

        'Parse last word
        If intSep1 = 0 Then intSep1 = Len(txtFrom) + 1

        'Parse word
        strWord = Mid$(txtFrom, i, intSep1 - i)

        strWord = ScrambleWord(strWord)

        'Put scrambled word in our temp string
        Mid$(strTemp, i, intSep1 - i) = strWord
        i = intSep1

        'Search for the next letter
        Do
            If i > Len(txtFrom) Then Exit Do
            If Asc(Mid$(txtFrom, i, 1)) > 64 Then Exit Do
            i = i + 1
        Loop

        i = i - 1

    Next i

    txtTo = strTemp

    cmdConvert.Enabled = True

End Sub

Private Function ScrambleWord(strWord As String) As String

  Dim arrWord() As Byte
  Dim strTemp As String
  Dim tmpItem As Variant
  Dim intWordLength As Integer

    strTemp = strWord

    If Len(strWord) > 3 Then

        ReDim arrWord(Len(strWord) - 1) As Byte

        arrWord(0) = CByte(Asc(Left$(strWord, 1)))
        arrWord(Len(strWord) - 1) = CByte(Asc(Right$(strWord, 1)))

        intWordLength = Len(strWord)
        strWord = Mid$(strWord, 2, Len(strWord) - 2)

        Do
            'Get random position
            Randomize
            intRandomID = CInt((intWordLength - 3) * Rnd) + 1

            If arrWord(intRandomID) = 0 Then
                arrWord(intRandomID) = CByte(Asc(Left$(strWord, 1)))  'Put left most letter in a random place in the array

                If Len(strWord) > 1 Then 'Remove left most letter
                    strWord = Mid$(strWord, 2)
                  Else 'Clear variable to end
                    strWord = ""
                End If
            End If

        Loop Until strWord = ""

        'Convert byte array to string
        strWord = StrConv(arrWord, vbUnicode) 'strWord & tmpItem

    End If

    'Return
    ScrambleWord = strWord

End Function

':) Ulli's VB Code Formatter V2.13.6 (09/23/2003 22:25:35) 0 + 115 = 115 Lines
