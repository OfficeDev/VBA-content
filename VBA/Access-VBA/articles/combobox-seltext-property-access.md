---
title: ComboBox.SelText Property (Access)
keywords: vbaac10.chm11437
f1_keywords:
- vbaac10.chm11437
ms.prod: access
api_name:
- Access.ComboBox.SelText
ms.assetid: dc2b46d7-c688-c9b5-c44f-c490a91589fe
ms.date: 06/08/2017
---


# ComboBox.SelText Property (Access)

The  **SelText** property returns a string containing the selected text. Read/write **String**.


## Syntax

 _expression_. **SelText**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

 If no text is selected, the **SelText** property contains a **Null** value.

The  **SelText** property uses a string expression that contains the text selected in the control. If the control contains selected text when this property is set, the selected text is replaced by the new **SelText** setting.

To set or return this property for a control, the control must have the focus. To move the focus to a control, use the  **SetFocus** method.


## Example

The following example uses two event procedures to search for text specified by a user. The text to search is set in the form's Load event procedure. The Click event procedure for the Find button (which the user clicks to start the search) prompts the user for the text to search for and selects the text in the text box if the search is successful.


```vb
Sub Form_Load() 
 Dim ctlTextToSearch As Control 
 Set ctlTextToSearch = Forms!Form1!TextBox1 
 ctlTextToSearch.SetFocus ' SetFocus to text box. 
 ctlTextToSearch.SelText = "This company places large orders " _ 
 &; "twice a year for garlic, oregano, chilies and cumin." 
End Sub 
 
Sub Find_Click() 
 Dim strSearch As String, intWhere As Integer 
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

