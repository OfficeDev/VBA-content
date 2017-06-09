---
title: GetFormat, GetText, SetText Methods Example
keywords: fm20.chm5225166
f1_keywords:
- fm20.chm5225166
ms.prod: office
ms.assetid: b17140cb-ab27-0073-8d7f-47eb91e31364
ms.date: 06/08/2017
---


# GetFormat, GetText, SetText Methods Example

The following example uses the  **GetFormat**, **GetText**, and **SetText** methods to transfer text between a **DataObject** and the Clipboard.

The user types text into a  **TextBox** and then can transfer it to a **DataObject** in a standard text format by clicking CommandButton1. Clicking CommandButton2 retrieves the text from the **DataObject**. Clicking CommandButton3 copies text from TextBox1 to the **DataObject** in a custom format. Clicking CommandButton4 retrieves the text from the **DataObject** in a custom format.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:



- A  **TextBox** named TextBox1.
    
- Four  **CommandButton** controls named CommandButton1 through CommandButton4.
    
- A  **Label** named Label1.
    




```vb
Dim MyDataObject As DataObject 
 
Private Sub CommandButton1_Click() 
'Put standard format on Clipboard 
 If TextBox1.TextLength > 0 Then 
 Set MyDataObject = New DataObject 
 MyDataObject.SetText TextBox1.Text 
 Label1.Caption = "Put on D.O." 
 CommandButton2.Enabled = True 
 CommandButton4.Enabled = False 
 End If 
End Sub 
 
Private Sub CommandButton2_Click() 
'Get standard format from Clipboard 
 If MyDataObject.GetFormat(1) = True Then 
 Label1.Caption = "Std format - " _ 
 &; MyDataObject.GetText(1) 
 End If 
End Sub 
 
Private Sub CommandButton3_Click() 
'Put custom format on Clipboard 
 If TextBox1.TextLength > 0 Then 
 Set MyDataObject = New DataObject 
 MyDataObject.SetText TextBox1.Text, 233 
 Label1.Caption = "Custom on D.O." 
 CommandButton4.Enabled = True 
 CommandButton2.Enabled = False 
 End If 
End Sub 
 
Private Sub CommandButton4_Click() 
'Get custom format from Clipboard 
 If MyDataObject.GetFormat(233) = True Then 
 Label1.Caption = "Cust format - " _ 
 &; MyDataObject.GetText(233) 
End If 
End Sub 
 
Private Sub UserForm_Initialize() 
 CommandButton2.Enabled = False 
 CommandButton4.Enabled = False 
End Sub
```


