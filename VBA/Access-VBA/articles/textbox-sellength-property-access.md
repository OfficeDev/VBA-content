---
title: TextBox.SelLength Property (Access)
keywords: vbaac10.chm11109
f1_keywords:
- vbaac10.chm11109
ms.prod: access
api_name:
- Access.TextBox.SelLength
ms.assetid: 0fb2371d-0f60-b0c7-5c4b-7a0689867b21
ms.date: 06/08/2017
---


# TextBox.SelLength Property (Access)

The  **SelLength** property specifies or determines the number of characters selected in a text box. Read/write **Integer**.


## Syntax

 _expression_. **SelLength**

 _expression_ A variable that represents a **TextBox** object.


## Remarks

The  **SelLength** property uses an **Integer** in the range 0 to the total number of characters in a text box or text box portion of a combo box.

To set or return this property for a control, the control must have the focus. To move the focus to a control, use the  **SetFocus** method.

Setting the  **SelLength** property to a number less than 0 produces a run-time error.


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


[TextBox Object](textbox-object-access.md)

