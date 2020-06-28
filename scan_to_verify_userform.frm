VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} scan_to_verify_userform 
   Caption         =   "Scan To Verify"
   ClientHeight    =   5685
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   5205
   OleObjectBlob   =   "scan_to_verify_userform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "scan_to_verify_userform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
        
    'Will place the current inv_bid value from sheck_check_userform in the current Userform's inv_bid textbox
    inv_bid_tb2.Value = shelf_check_userform.inv_bid_tb.Value
    cart_num = "N/a"
    shelf_num = "N/a"
    
    'Activates "Shelf_Check" Sheet, then adds a "Verified" column with bold text and a green backround.
    Sheets("Shelf_Check").Activate
    
    Dim verified_header As Range: Set verified_header = Range("E1")
    
    verified_header.Value = "Verified"
    verified_header.Interior.ColorIndex = 4
    Worksheets("Shelf_Check").Range("E:E").Font.Bold = True
    
    
    emptyRow = WorksheetFunction.CountA(Range("A:A")) + 1
    
    'Loops through each non-empty row, if the inv_bid (copied from the shelf_check_useform) is already on the sheet, show the
    'cart and shelf location on the current userform via labels: cart_num & shelf_num.
    For i = 1 To emptyRow - 1
        
            inv_bid_i = Cells(i, 3).Value
            cart_i = Cells(i, 1).Value
            shelf_i = Cells(i, 2).Value
            
            If CStr(inv_bid_i) = CStr(inv_bid_tb2.Value) Then
            
                num = i
                bid = inv_bid_tb2.Value
                cart_num = cart_i
                shelf_num = shelf_i
                
            End If
        
    Next i
    
    'Ensures the "Match" textbox is blank and the user's cursor is placed in said textbox.
    match_tb.Value = ""
    match_tb.SetFocus
    
End Sub

'Evaluates each button press within the "Match" textbox. If the button pressed is Enter, return a Right button press instead.
'This ensures that the user's cursor does not leave the current text box after pressing Enter.
'Once Enter is pressed, execute the code found in the "verify_btn_Click" subroutine.
Private Sub match_tb_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
         KeyCode = vbKeyRight
         Call verify_btn_Click
    End If

End Sub


Private Sub verify_btn_Click()
            
            'If the values of both textboxes is equal...
            If CStr(inv_bid_tb2.Value) = CStr(match_tb.Value) Then
            
                'change the userform's color to green and its text to black.
                scan_to_verify_userform.BackColor = RGB(124, 252, 0) ' green
                Label1.ForeColor = lngBlack
                Label2.ForeColor = lngBlack
                
                emptyRow = WorksheetFunction.CountA(Range("A:A")) + 1
                
                'Loops through each row, if the value placed in the textboxes is already on the Sheet...
                For i = 1 To emptyRow - 1
        
                    inv_bid_i = Cells(i, 3).Value
            
                    If CStr(inv_bid_i) = CStr(match_tb.Value) Then
                    
                        'add the word "True" with a green backround next to the matching inv_bid's Cell/row.
                        Cells(i, 5).Value = "True"
                        Cells(i, 5).Interior.ColorIndex = 4
                    
                    End If
        
                Next i
                

                'Play sound effect #3 then clear Match textbox's text and place user cursor in said textbox.
                Call PlayWAV3
                
                match_tb.Value = ""
                match_tb.SetFocus
                
            'If the two textboxes do not have the same value, do the opposite of the block above.
            Else
                scan_to_verify_userform.BackColor = RGB(255, 0, 0) ' red
                Label1.ForeColor = lngBlack
                Label2.ForeColor = lngBlack
                
                emptyRow = WorksheetFunction.CountA(Range("A:A")) + 1
                
                For i = 1 To emptyRow - 1
        
                    inv_bid_i = Cells(i, 3).Value
            
                    If CStr(inv_bid_i) = CStr(inv_bid_tb2.Value) Then
                        
                        Cells(i, 5).Value = "False"
                        Cells(i, 5).Interior.ColorIndex = 3
                    
                    End If
        
                Next i
                
                'Sound effect #1 with message to the user informing them of mismatch. Newlines were added for readability.
                Call PlayWAV
                MsgBox "Inv BIDs do not match:" & vbNewLine & vbNewLine & inv_bid_tb2.Value & vbNewLine & vbNewLine & match_tb.Value
                
                match_tb.Value = ""
                match_tb.SetFocus
                
            End If
    
End Sub

'Resets all values to blank and places user's cursor in the inv_bid textbox
Private Sub reset_btn_Click()
    inv_bid_tb2.Value = ""
    match_tb.Value = ""
    cart_num = ""
    shelf_num = ""
    inv_bid_tb2.SetFocus

End Sub

'Closes scan_to_verify userform.
Private Sub complete_btn_Click()

    Unload Me
    
End Sub
