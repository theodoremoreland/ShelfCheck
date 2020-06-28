Attribute VB_Name = "Module1"
'Following conditional statements will execute based on the version of Excel that is being used.
#If VBA7 Then

  Private Declare PtrSafe Function PlaySound Lib "winmm.dll" _
  Alias "PlaySoundA" (ByVal lpszName As String, _
  ByVal hModule As Long, ByVal dwFlags As Long) As Long
  
#Else
  
  Private Declare PtrSafe Function PlaySound Lib "winmm.dll" _
  Alias "PlaySoundA" (ByVal lpszName As String, _
  ByVal hModule As Long, ByVal dwFlags As Long) As Long
  
#End If
 
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_FILENAME = &H20000
 
'Creates a sound effect by binding a '.wav' file found in the given directory.
Sub PlayWAV()
    Dim WAVFile As String
    WAVFile = "C:\Users\GKQNG\Desktop\siren.wav"
    Call PlaySound(WAVFile, 0&, SND_ASYNC Or SND_FILENAME)
End Sub

'Sound effect #2.
Sub PlayWAV2()
    Dim WAVFile2 As String
    WAVFile2 = "C:\Windows\Media\chord.wav"
    Call PlaySound(WAVFile2, 0&, SND_ASYNC Or SND_FILENAME)
End Sub

'Sound effect #3.
Sub PlayWAV3()
    Dim WAVFile2 As String
    WAVFile2 = "C:\Windows\Media\tada.wav"
    Call PlaySound(WAVFile2, 0&, SND_ASYNC Or SND_FILENAME)
End Sub



Sub Shelf_Check()
    
    'Loops through each Sheet. If no Sheet named "Shelf_Check" is found, creates one.
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = "Shelf_Check" Then
            exists = True
        End If
    Next i

    If Not exists Then
        Sheets.Add
        ActiveSheet.Name = "Shelf_Check"
    End If

    'Creates variables with the location for each header.
    Dim cart_header As Range: Set cart_header = Range("A1")
    Dim shelf_header As Range: Set shelf_header = Range("B1")
    Dim inv_bid_header As Range: Set inv_bid_header = Range("C1")
    Dim scans_header As Range: Set scans_header = Range("D1")
    
    'Adds text to headers.
    cart_header.Value = "Cart #"
    shelf_header.Value = "Shelf #"
    inv_bid_header.Value = "Inv_BID"
    scans_header.Value = "Scans"
    
    'Makes the headers yellow.
    cart_header.Interior.ColorIndex = 6
    shelf_header.Interior.ColorIndex = 6
    inv_bid_header.Interior.ColorIndex = 6
    scans_header.Interior.ColorIndex = 6

    'Formats columns 1, 2, and 4 for bold text.
    Worksheets("Shelf_Check").Range("A:A").Font.Bold = True
    Worksheets("Shelf_Check").Range("B:B").Font.Bold = True
    Worksheets("Shelf_Check").Range("D:D").Font.Bold = True
    
    'Opens the sheck_check userform.
    shelf_check_userform.Show


End Sub


