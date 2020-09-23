VERSION 5.00
Begin VB.Form mainwindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "word frequency analyser"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   Icon            =   "word_freq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8010
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSort 
      Height          =   330
      Left            =   2700
      TabIndex        =   5
      Top             =   4545
      Width           =   2040
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   7815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e&xit"
      Height          =   330
      Left            =   6780
      TabIndex        =   2
      Top             =   4545
      Width           =   1155
   End
   Begin VB.CommandButton parsebutton 
      Caption         =   "Parse"
      Height          =   330
      Left            =   120
      TabIndex        =   1
      Top             =   4545
      Width           =   1425
   End
   Begin VB.TextBox main 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      HideSelection   =   0   'False
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "word_freq.frx":0442
      Top             =   480
      Width           =   7815
   End
   Begin VB.Label Label1 
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   135
      TabIndex        =   4
      Top             =   90
      Width           =   7710
   End
End
Attribute VB_Name = "mainwindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const reset_Text = "Double click in this textbox to paste contents of clipboard to this window."
Private strInput As String          'this will hold what's copied to the textbox
Private bytInput() As Byte          'we'll put the string into this array
Private iCursor As Long             'to keep our place in bytInput()
Private bytOutput() As Byte         'we'll use this array to hold valid words
Private jCursor As Long             'to keep our place in bytOutput()
Private pbyte As Byte               'the byte we're currently working on
Private byteCount As Byte           'used to track length of words in bytInput()
Private bytWords() As Byte          'two dimensional array - to hold words
Private theWords() As Byte          'two dim array - holds words
Private intHist() As Integer        'store frequency of words
Private wordIndex() As Long         'where a word starts in one dimension array
Private wordLen() As Long           'length of word in one dimension array
Private theWordLen() As Long        'length of word in new one dimension array
Private theWordIndex() As Long      'where a word starts in new one dimension array
Private theWordCount As Long        'total of words in document
Private bytArrayLen As Long         'length of one dimension array
Private longestWord As Byte         'length of longest word found
Private theHist() As Long           'number of occurances of a word
Private sortKey As String
Private startTime As Long           'used to time procedures
Private Declare Function timeGetTime Lib "winmm.dll" () As Long         'used to time procedures

Private Sub cmdSort_Click()

    displayResults
    
End Sub

Private Sub Command1_Click()

    Erase theWords
    Erase theWordLen
    Erase theWordIndex
    Erase theHist
    End
    
End Sub

Private Sub Form_Load()
    main.Text = reset_Text
    List1.Visible = False
    Label1.Visible = False
    cmdSort.Visible = False
End Sub

Private Sub main_DblClick()
    main.Text = Clipboard.GetText
End Sub

Private Sub parsebutton_Click()
    parsebutton.Visible = False
    strInput = main.Text
    Call mainPath
End Sub

Private Sub Resetbutton_Click()
    Call reset
End Sub

Private Sub mainPath()
    
    Dim i As Long
    Dim j As Long
    
    If strInput = "" Then
        MsgBox "No Words found"
        Exit Sub
    End If
    'put restart code here
    
    startTime = timeGetTime
    
    'let's put the ANSI string into an array of bytes
    'it's faster than string operations
    ReDim bytInput(1 To Len(strInput))
    ReDim bytOutput(1 To Len(strInput))
    
    'put file characters into array
    For i = 1 To Len(strInput)
        bytInput(i) = Asc(Mid$(strInput, i, 1))
    Next i
    
    'loop through the file and parse - eliminate 1 and 2 character words
    byteCount = 1       'used to eliminate small words
    jCursor = 1
    For iCursor = 1 To UBound(bytInput)
        pbyte = bytInput(iCursor)
        parseTheFile
    Next iCursor
    
    shapeBytArray
'    showBytOutputInDebugWindow           'debugging purposes
    make2DArray
    markDupesMakeHist
    shuffleUp
'    sortKey = "none"
'        sortKey = "hist"
'                sortKey = "length"
    cmdSort.Visible = True
    displayResults
    Debug.Print timeGetTime - startTime & " msec"
    
End Sub
Private Sub parseTheFile()

    'bytes that are valid characters (letters and the single quote)
    'are put into the bytOutput array and the cursor is incremented
    'other characters are converted to chr(32) - space characters.
    'if a space is already at the end of the array, and the next byte
    'to be added is also a space, it is discarded.
    'if a word of 1 or 2 bytes has been added, and a space would be added
    'next, the cursor is moved back to the space before the short word
    'and the space is discarded.
    

    Select Case pbyte
    
        Case 65 To 90                       'A to Z
            pbyte = pbyte + 32              'convert to lowercase
            bytOutput(jCursor) = pbyte      'add character to output array
            jCursor = jCursor + 1           'increment cursor
            byteCount = byteCount + 1       'count the characters in current word
            
            
        Case 97 To 122, 39                  'a to z,  and single quote
            bytOutput(jCursor) = pbyte      'add character to output array
            jCursor = jCursor + 1           'increment cursor
            byteCount = byteCount + 1       'count characters in current word
            
            
        Case Else
                If bytOutput(jCursor) <> 32 Then
                    If byteCount > 2 Then               'the previous word has at least 3 bytes
                        bytOutput(jCursor) = 32         'append a space
                        jCursor = jCursor + 1           'increment cursor
                    Else
                        jCursor = jCursor - byteCount   'move cursor back to previous word
                        If jCursor < 1 Then             'boundary check - are we at the beginning of the array
                            jCursor = 1                 'output array must start at 1
                        End If
                    End If
                End If
            byteCount = 0                               'reset to calculate next word length

    End Select

End Sub
Private Sub make2DArray()

    'put parsed 1 dimension array - bytOutput()
    'into 2 dimension jagged array - bytWords()
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    'find longest word to redim 2nd index of bytWords()
    longestWord = 0
    For i = 1 To UBound(wordLen)
        If wordLen(i) > longestWord Then longestWord = wordLen(i)
    Next
    
    ReDim bytWords(1 To theWordCount, 1 To longestWord)
    
    For i = 1 To theWordCount
        For j = wordIndex(i) To wordIndex(i) + wordLen(i) - 1
            k = k + 1
            bytWords(i, k) = bytOutput(j)
        Next j
        k = 0
    Next i
    
End Sub

Public Sub markDupesMakeHist()

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim thisWord() As Byte

    ReDim intHist(1 To theWordCount)    'extreme case - all words are the same

    ReDim thisWord(longestWord)

    For i = 1 To theWordCount

        'set thisWord(j) = byteArray(i,x)
        For j = 1 To wordLen(i)
            thisWord(j) = bytWords(i, j)
'            Debug.Print Chr$(thisWord(j));
        Next j
'        Debug.Print
        For k = i + 1 To theWordCount
            If bytWords(k, 1) <> 0 Then     'don't test if marked dupe
                'compare first bytes
                If thisWord(1) = bytWords(k, 1) Then     'maybe a match
                    If allBytesMatch(thisWord, bytWords, k) Then
                        bytWords(k, 1) = 0     'mark as duplicate
                        intHist(i) = intHist(i) + 1
                    End If
                End If
            End If
        Next k
    Next i
    
End Sub
Private Function allBytesMatch(firstArray() As Byte, secondArray() As Byte, firstDimIndex As Integer) As Boolean

    Dim i As Integer
    
    For i = 1 To wordLen(firstDimIndex)
        If firstArray(i) <> secondArray(firstDimIndex, i) Then
            allBytesMatch = False
            Exit Function
        End If
    Next i
    allBytesMatch = True
    
End Function
Private Sub displayResults()

    Dim i As Long
    Dim j As Long
    Dim strOut As String
    Dim theCaption As String
    Dim myTab As String
    myTab = Space(4)
    
    main.Visible = False
    List1.Visible = True
    Label1.Visible = True
    
    List1.Clear
    
    Select Case cmdSort.Caption
    
        Case "Sort By Frequency"
        
            theCaption = "Frequency" & myTab & "Length" & myTab & "Word"
            Label1.Caption = theCaption
            
            For i = 1 To theWordCount
                strOut = ""
                strOut = Right$("000" & theHist(i), 4)
                strOut = strOut & vbTab & vbTab & theWordLen(i) & vbTab
                For j = 1 To theWordLen(i)
                    strOut = strOut & Chr$(theWords(i, j))
                Next j
                List1.AddItem strOut
            Next i
            cmdSort.Caption = "Sort by Length"
        
        Case "Sort by Length"
        
            theCaption = "Length" & myTab & "Frequency" & myTab & "Word"
            Label1.Caption = theCaption
            
            For i = 1 To theWordCount
                strOut = ""
                strOut = Right$("000" & theWordLen(i), 4)
                strOut = strOut & vbTab & vbTab & theHist(i) & vbTab
                For j = 1 To theWordLen(i)
                    strOut = strOut & Chr$(theWords(i, j))
                Next j
                List1.AddItem strOut
            Next i
            cmdSort.Caption = "Original Order"
        
        
        Case Else
    
            theCaption = "Word" & Space(longestWord)
            theCaption = Left$(theCaption, longestWord + 3) & "Length" & myTab & myTab & "Frequency"
            Label1.Caption = theCaption
            
            For i = 1 To theWordCount
                strOut = ""
                For j = 1 To theWordLen(i)
                    strOut = strOut & Chr$(theWords(i, j))
                Next j
                strOut = strOut & Space(longestWord)
                strOut = Left$(strOut, longestWord + 3) & vbTab & theWordLen(i) & vbTab & vbTab & theHist(i)
                List1.AddItem strOut
            Next i
            cmdSort.Caption = "Sort By Frequency"
            
        End Select
    
End Sub


Private Sub shuffleUp()

    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    'these new arrays will hold non-duplicated values
    
    ReDim theWords(1 To theWordCount, 1 To longestWord)
    ReDim theWordLen(1 To theWordCount)
    ReDim theWordIndex(1 To theWordCount)
    ReDim theHist(1 To theWordCount)
    
    'copy non-duplicated values to new array
    
    k = 1                                           'init cursor to new array
    For i = 1 To theWordCount
        If bytWords(i, 1) <> 0 Then                 'not marked as dupe, so...
            For j = 1 To wordLen(i)                 'move to new array
                theWords(k, j) = bytWords(i, j)
            Next j
            theWordLen(k) = wordLen(i)              'move to new array
            theWordIndex(k) = wordIndex(i)          'move to new array
            theHist(k) = intHist(i) + 1             'move to new array
            k = k + 1                               'increment cursor to new array
        End If
    Next i

    theWordCount = k - 1                            'reset to non-dupe value
    
    Erase bytInput
    Erase bytOutput
    Erase bytWords
    Erase intHist
    Erase wordIndex
    Erase wordLen


    
'    ShowNewArraysInDebugWindow
    
End Sub
Private Sub ShowNewArraysInDebugWindow()

    Dim i As Long
    Dim j As Long
    
    Debug.Print "the new arrays"
    Debug.Print "#", "theWordIndex", "theWordLen", "theHist", "theWords"
    For i = 1 To theWordCount
        Debug.Print i, theWordIndex(i), theWordLen(i), theHist(i),
        For j = 1 To theWordLen(i)
            Debug.Print Chr$(theWords(i, j));
        Next j
        Debug.Print
    Next i
    Debug.Print "end of new arrays"

End Sub

Private Sub showBytOutputInDebugWindow()

    Dim i As Long
    
    Debug.Print "word #", "wordIndex", "wordLen", "word", "bytes"
    For i = 1 To theWordCount
        Debug.Print i, wordIndex(i), wordLen(i),
        showWordandBytes (i)
    Next
    
End Sub
Private Sub shapeBytArray()

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim lastByt As Long


    lastByt = UBound(bytOutput)
    If bytOutput(lastByt) <> 0 Then
        ReDim Preserve bytOutput(1 To lastByt + 2)
        bytOutput(lastByt + 1) = 0
        bytOutput(lastByt + 2) = 0
    End If
    
    
    'ensure that the last three bytes of the array are - a valid character, 32, then 0
    For i = UBound(bytOutput) To 1 Step -1
        If bytOutput(i) <> 0 Then           'found first non-zero byte
            If bytOutput(i) <> 32 Then      'if last valid char is not already a space
                bytOutput(i + 1) = 32       'then make the character after it a space
            End If
            Exit For
        End If
    Next i
    
    bytArrayLen = i + 2
    
    'all words end in a space
    'how many spaces are there (how many words are there)
    For i = 1 To bytArrayLen
        If bytOutput(i) = 32 Then
            theWordCount = theWordCount + 1
        End If
    Next i
    
'    Debug.Print theWordCount & " words"
    
    'where are the spaces?
    ReDim wordIndex(1 To theWordCount)
    ReDim wordLen(1 To theWordCount)
    
    wordIndex(1) = 1                    'first word starts at start of array
    j = 2
    For i = 1 To bytArrayLen
        If bytOutput(i) = 32 Then
            If bytOutput(i + 1) <> 0 Then   'end of array is after a space
                wordIndex(j) = i + 1        'beginning of next word is after a space
                j = j + 1                   'increment wordIndex index
            End If
        End If
    Next
    
    If theWordCount > 1 Then

        'every word but the first is enclosed in spaces
        For j = 2 To theWordCount - 1
            wordLen(j) = wordIndex(j + 1) - wordIndex(j) - 1
        Next j
        
        'first word
        wordLen(1) = wordIndex(2) - 2
        
        'last word
        i = wordIndex(theWordCount)
        Do Until bytOutput(i) = 32
            wordLen(theWordCount) = wordLen(theWordCount) + 1
            i = i + 1
        Loop
    Else
        'trivial case - one word
        wordIndex(1) = 1
        wordLen(1) = bytArrayLen - 2
    End If

End Sub

Private Sub showWordandBytes(Value As Long)

    Dim i As Long

    For i = wordIndex(Value) To wordIndex(Value) + wordLen(Value) - 1   'don't include space char
        Debug.Print Chr$(bytOutput(i));
    Next
    Debug.Print " - ";
    For i = wordIndex(Value) To wordIndex(Value) + wordLen(Value) - 1   'don't include space char
        Debug.Print (bytOutput(i));
    Next
    Debug.Print
    
End Sub
