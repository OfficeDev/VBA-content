---
title: Application.Explorers Property (Outlook)
keywords: vbaol11.chm720
f1_keywords:
- vbaol11.chm720
ms.prod: outlook
api_name:
- Outlook.Application.Explorers
ms.assetid: bbbdbd6e-a238-8108-fbbd-5f7d7821aaa7
ms.date: 06/08/2017
---


# Application.Explorers Property (Outlook)

Returns an  **[Explorers](explorers-object-outlook.md)** collection object that contains the **[Explorer](explorer-object-outlook.md)** objects representing all open explorers. Read-only.


## Syntax

 _expression_ . **Explorers**

 _expression_ A variable that represents an **Application** object.


## Example

The following Microsoft Visual Basic for Applications (VBA) example displays the number of explorer windows that are open.


```vb
Private Sub CountExplorers() 
 
 MsgBox "There are " &; _ 
 
 Application.Explorers.Count &; " Explorers." 
 
End Sub
```

The following VBA example uses the  **[Count](selection-count-property-outlook.md)** property and **[Item](selection-item-method-outlook.md)** method of the **[Selection](selection-object-outlook.md)** collection returned by the **Selection** property to display the senders of all mail items selected in the explorer that displays the **Inbox**. To run this example, you need to have at least one mail item selected in the explorer displaying the Inbox. You might receive an error if you select items other than a mail item such as task request as the  **SenderName** property does not exist for a **[TaskRequestItem](taskrequestitem-object-outlook.md)** object.




```vb
Sub GetSelectedItems() 
 
 Dim myOlExp As Outlook.Explorer 
 
 Dim myOlSel As Outlook.Selection 
 
 Dim MsgTxt As String 
 
 Dim x As Integer 
 
 
 
 MsgTxt = "You have selected items from: " 
 
 Set myOlExp = Application.Explorers.Item(1) 
 
 If myOlExp = "Inbox" Then 
 
 Set myOlSel = myOlExp.Selection 
 
 For x = 1 To myOlSel.Count 
 
 MsgTxt = MsgTxt &; myOlSel.Item(x).SenderName &; ";" 
 
 Next x 
 
 MsgBox MsgTxt 
 
End If 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-outlook.md)

