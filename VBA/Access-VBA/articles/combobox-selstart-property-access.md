---
title: ComboBox.SelStart Property (Access)
keywords: vbaac10.chm11438
f1_keywords:
- vbaac10.chm11438
ms.prod: access
api_name:
- Access.ComboBox.SelStart
ms.assetid: 056196b5-828a-f276-da26-983c8b47cd05
ms.date: 06/08/2017
---


# ComboBox.SelStart Property (Access)

The  **SelStart** property specifies or determines the starting point of the selected text or the position of the insertion point if no text is selected. Read/write **Integer**.


## Syntax

 _expression_. **SelStart**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

The  **SelStart** property uses an **Integer** in the range 0 to the total number of characters in the text box portion of a combo box.

To set or return this property for a control, the control must have the focus. To move the focus to a control, use the  **SetFocus** method.

Changing the  **SelStart** property cancels the selection, places an insertion point in the text, and sets the **SelLength** property to 0.


## Example

The following example uses two event procedures to search for text specified by a user. The text to search is set in the form's Load event procedure. The Click event procedure for the Find button (which the user clicks to start the search) prompts the user for the text to search for and selects the text in the text box if the search is successful.


```vb
Private Sub Form_Load() 
 
 Dim ctlTextToSearch As Control 
 Set ctlTextToSearch = Forms!Form1!Textbox1 
 
 ' SetFocus to text box. 
 ctlTextToSearch.SetFocus 
 ctlTextToSearch.Text = "This company places large orders twice " &; _ 
 "a year for garlic, oregano, chilies and cumin." 
 Set ctlTextToSearch = Nothing 
 
End Sub 
 
Public Sub Find_Click() 
 
 Dim strSearch As String 
 Dim intWhere As Integer 
 Dim ctlTextToSearch As Control 
 
 ' Get search string from user. 
 With Me!Textbox1 
 strSearch = InputBox("Enter text to find:") 
 
 ' Find string in text. 
 intWhere = InStr(.Value, strSearch) 
 If intWhere Then 
 ' If found. 
 .SetFocus 
 .SelStart = intWhere - 1 
 .SelLength = Len(strSearch) 
 Else 
 ' Notify user. 
 MsgBox "String not found." 
 End If 
 End With 
 
End Sub
```


## See also


#### Concepts


[ComboBox Object](combobox-object-access.md)

