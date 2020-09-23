<div align="center">

## A ComboBox Procedure to save its contents\.


</div>

### Description

A simple procedure you can use it works like the Internet Explorer Address Bar. It will save what is typed into the Combo Box and it does not allow duplicate entries to be entered. Please vote if you like it. Sendmessage API can be used also for speed. This is a little simpler because there no declarations.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[thasnake](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/thasnake.md)
**Level**          |Intermediate
**User Rating**    |4.0 (12 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/thasnake-a-combobox-procedure-to-save-its-contents__1-41733/archive/master.zip)





### Source Code

```
Private Sub ComboSave(cboName As ComboBox)
On Error GoTo Err:
Dim val As String
Dim i As Long
Dim match As Boolean
val = Trim(cboName.Text)
For i = 0 To cboName.ListCount - 1
 If cboName.List(i) = val Then
  match = True
  Exit For
 Else
  match = False
 End If
Next i
If match = False Then
 cboName.AddItem val
 'You could add code here to save the values
 'to a file or registry or something like that
 'so they can be loaded back in next time
 'program is started
End If
Exit Sub
Err:
MsgBox "Sorry, there was an error!. " & vbCrLf & _
  "Please try again.", vbExclamation, "Error"
End Sub
```

