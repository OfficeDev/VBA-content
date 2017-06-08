---
title: Application.Inspectors Property (Outlook)
keywords: vbaol11.chm721
f1_keywords:
- vbaol11.chm721
ms.prod: outlook
api_name:
- Outlook.Application.Inspectors
ms.assetid: c2dde847-d033-90e3-30d2-62ff375d6843
ms.date: 06/08/2017
---


# Application.Inspectors Property (Outlook)

Returns an  **[Inspectors](inspectors-object-outlook.md)** collection object that contains the **[Inspector](inspector-object-outlook.md)** objects representing all open inspectors. Read-only.


## Syntax

 _expression_ . **Inspectors**

 _expression_ A variable that represents an **Application** object.


## Example

This Microsoft Visual Basic example uses the  **[Inspectors](application-inspectors-property-outlook.md)** property and the **[Count](inspectors-count-property-outlook.md)** property and **[Item](inspectors-item-method-outlook.md)** method of the **[Inspectors](inspectors-object-outlook.md)** object to display the captions of all inspector windows.


```vb
Private Sub CommandButton1_Click() 
 
 Dim myInspectors As Outlook.Inspectors 
 
 Dim x as Integer 
 
 Dim iCount As Integer 
 
 
 
 Set myInspectors = Application.Inspectors 
 
 iCount = Application.Inspectors.Count 
 
 If iCount > 0 Then 
 
 For x = 1 To iCount 
 
 MsgBox myInspectors.Item(x).Caption 
 
 Next x 
 
 Else 
 
 MsgBox "No inspector windows are open." 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-outlook.md)

