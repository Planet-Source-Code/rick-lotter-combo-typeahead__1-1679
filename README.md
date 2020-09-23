<div align="center">

## Combo\_TypeAhead


</div>

### Description

Allows the VB combobox to have a type-ahead feature like the combobox in Access.

If there are any matching items in the combobox list, it will automatically

"fill in" the missing portions of the item and select it.

This will work for multiple characters, not just the first character of the string.

*** This is an expanded version of the original code ***

The updates allow for new items to be added into the list, automatically. It also

will handle the delete key (the previous code did not.)
 
### More Info
 
The combobox that you want to scan for entries in its list.

Optionally, you can specify whether the comparison is case-sensitive.

The function should be called as following (assuming a combobox call cboMine):

Private Sub cboMine_Change()

'If the last key was not a special key (control), then use typeahead function

If intLastKey >= 32 Then Call Combo_TypeAhead(cboMine)

End Sub

Private Sub cboMine_KeyDown(KeyCode As Integer, Shift As Integer)

'If the last key was a delete, then send a backspace to clear the selection

If KeyCode = vbKeyDelete Then SendKeys "{BACKSPACE}", True

End Sub

Private Sub cboMine_KeyPress(KeyAscii As Integer)

'Sets the last key value for control character checking

intLastKey = KeyAscii

End Sub

Private Sub cboMine_LostFocus()

Call Combo_AddNew(cboMine)

End Sub


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rick Lotter](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rick-lotter.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rick-lotter-combo-typeahead__1-1679/archive/master.zip)

### API Declarations

Private intLastKey As Integer ' Records last keypress


### Source Code

```
Public Sub Combo_AddNew(ByRef cboCurrent As ComboBox, _
  Optional blnCaseSensitive As Boolean = False, _
  Optional blnAddAsUpperCase As Boolean = True)
Dim lngServerNum As Long
Dim blnFoundMatch As Boolean
Dim strNewItem As String, strCurrentItem As String
strNewItem = cboCurrent.Text
If Not blnCaseSensitive Then strNewItem = UCase(strNewItem)
'Search for matches
blnFoundMatch = False
For lngServerNum = 0 To cboCurrent.ListCount - 1
 strCurrentItem = cboCurrent.List(lngServerNum)
 If Not blnCaseSensitive Then strCurrentItem = UCase(strCurrentItem)
 If strCurrentItem = strNewItem Then blnFoundMatch = True
Next lngServerNum
'If one is found, add and re-select
If Not blnFoundMatch Then
 If Not blnAddAsUpperCase Then
  cboCurrent.AddItem cboCurrent.Text
 Else
  cboCurrent.AddItem UCase(cboCurrent.Text)
 End If
 cboCurrent.ListIndex = cboCurrent.NewIndex
End If
End Sub
Public Sub Combo_TypeAhead(ByRef cboCurrent As ComboBox, _
  Optional blnCaseSensitive As Boolean = False)
'This function will allow the combobox cboCurrent to have the type-ahead feature _
found in Access. When the user types in text, it will look for a matching item in the _
list and add the remainder of the item on, and highlight the text.
'By default, the comparison is not case sensitive. If blnCaseSensitive is overridden _
with a true value, then it will consider case in the comparison.
Dim lngItemNum As Long, lngSelectedLength As Long, lngMatchIndex As Long
Dim strSearchText As String, strCurrentText As String
'Check for empty control, and abort if found
If cboCurrent.Text = "" Then Exit Sub
'Set up initial values for search
lngMatchIndex = -1
strSearchText = cboCurrent.Text
If Not blnCaseSensitive Then strSearchText = UCase(strSearchText)
lngSelectedLength = Len(strSearchText)
'Search all items for first match
For lngItemNum = 0 To cboCurrent.ListCount - 1
 strCurrentText = Mid(cboCurrent.List(lngItemNum), 1, lngSelectedLength)
 If Not blnCaseSensitive Then strCurrentText = UCase(strCurrentText)
 'If a match is found, record it and abort loop
 If strSearchText = strCurrentText Then
  lngMatchIndex = lngItemNum
  Exit For
 End If
Next lngItemNum
'If a match was found, select it and highlight the "filled in" text
If lngMatchIndex >= 0 Then
 cboCurrent.ListIndex = lngMatchIndex
 cboCurrent.SelStart = lngSelectedLength
 cboCurrent.SelLength = Len(cboCurrent.List(cboCurrent.ListIndex)) - lngSelectedLength
End If
End Sub
```

