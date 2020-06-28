VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} shelf_check_userform 
   Caption         =   "Shelf Check"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5010
   OleObjectBlob   =   "shelf_check_userform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "shelf_check_userform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Makes each textbox blank upon opening userform.
Private Sub UserForm_Initialize()
          
    cart_tb.Value = ""
    
    shelf_tb.Value = ""
    
    inv_bid_tb.Value = ""
    
End Sub


'Evaluates each button press within the "Match" textbox. If the button pressed is Enter, return a Right button press instead.
'This ensures that the user's cursor does not leave the current text box after pressing Enter.
'Once Enter is pressed, execute the code found in the "continue_btn_Click" subroutine.
Private Sub inv_bid_tb_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
         KeyCode = vbKeyRight
         Call continue_btn_Click
    End If

End Sub


Private Sub continue_btn_Click()

    Dim emptyRow As Long

    Sheets("Shelf_Check").Activate
    
    emptyRow = WorksheetFunction.CountA(Range("A:A")) + 1
    
    Range("C:C").Select
    
    'Searches for value entered into inv bid textbox.
    Set inv_bid = Selection.Find(What:=inv_bid_tb.Value, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
        
    
    'If any of the textboxes are empty (after pressing Enter) in the inv_bod text box, show user a message.
    If (inv_bid_tb.Value = "") Or (shelf_tb.Value = "") Or (cart_tb.Value = "") Then
        MsgBox ("All fields must have a value.")
        
    'If inv bid entered into inv bid textbox is not found...
    ElseIf (inv_bid) Is Nothing Then
        
        'Change userform color to green and add values entered into each textbox to the first empty row.
        shelf_check_userform.BackColor = RGB(124, 252, 0) ' green
        
        Cells(emptyRow, 1).Value = cart_tb.Value
        Cells(emptyRow, 2).Value = shelf_tb.Value
        Cells(emptyRow, 3).Value = inv_bid_tb.Value
        Cells(emptyRow, 4).Value = 1
        
        inv_bid_tb.Value = ""
        inv_bid_tb.SetFocus
        
    'If inv bid is found...
    Else
        'Loop through every row and assign current cells a variable.
        For i = 1 To emptyRow - 1
        
            cart_i = Cells(i, 1).Value
            shelf_i = Cells(i, 2).Value
            inv_bid_i = Cells(i, 3).Value
            
            'Once the current cell finds the inv bid...
            If CStr(inv_bid_i) = CStr(inv_bid_tb.Value) Then
            
                'Create variables for the current row and its Cells.
                num = i
                bid = inv_bid_tb.Value
                cart = cart_i
                shelf = shelf_i
                
            End If
        
        Next i
        
        'Nested If: (If the inv bid is found (and the cart and shelf entered are the same...)...
        If CStr(cart) = CStr(cart_tb.Value) And CStr(shelf) = CStr(shelf_tb.Value) Then
            
            'Play sound effect #1
            Call PlayWAV
        
            'Change row/cells background to yellow and add 1 to the scans cell in current row.
            Cells(num, 1).Interior.ColorIndex = 6
            Cells(num, 2).Interior.ColorIndex = 6
            Cells(num, 3).Interior.ColorIndex = 6
            Cells(num, 4).Value = Cells(num, 4).Value + 1
            
            shelf_check_userform.BackColor = RGB(255, 255, 0) ' yellow
            
            'Reset cursor and inv bid text box.
            inv_bid_tb.Value = ""
            inv_bid_tb.SetFocus
        
        'Nested else: (If the inv bid is found (and the cart and shelf entered are not the same...)...
        Else
            'Play sound effect #2, change userform color to red, send user a message and open scan_to_verify userform.
            Call PlayWAV2
            shelf_check_userform.BackColor = RGB(255, 0, 0) ' red
            MsgBox "This Inv_BID (" & bid & ") has already been scanned @ Cart: " & cart & ", Shelf: " & shelf
            
            scan_to_verify_userform.Show
            
            inv_bid_tb.Value = ""
            inv_bid_tb.SetFocus
            
        End If
    End If
    
End Sub

'Opens scan_to_verify_userform
Private Sub verify_btn_Click()
    scan_to_verify_userform.Show
End Sub

'Closes userform.
Private Sub cancel_btn_Click()
    Unload Me
End Sub
