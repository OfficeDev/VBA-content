---
title: Application.ActiveExplorer Method (Outlook)
keywords: vbaol11.chm712
f1_keywords:
- vbaol11.chm712
ms.prod: outlook
api_name:
- Outlook.Application.ActiveExplorer
ms.assetid: f6dd27c0-4319-c7fc-191f-8b3b2ea319d3
ms.date: 06/08/2017
---


# Application.ActiveExplorer Method (Outlook)

Returns the topmost  **[Explorer](explorer-object-outlook.md)** object on the desktop.


## Syntax

 _expression_ . **ActiveExplorer**

 _expression_ A variable that represents an **Application** object.


### Return Value

An  **Explorer** that represents the topmost explorer on the desktop. Returns **Nothing** if no explorer is active.


## Remarks

 Use this method to return the **Explorer** object that the user is most likely viewing. This method is also useful for determining when there is no active explorer, so a new one can be opened.


## Example

The following Microsoft Visual Basic for Applications (VBA) example uses the  **[Count](selection-count-property-outlook.md)** property and **[Item](selection-item-method-outlook.md)** method of the **[Selection](selection-object-outlook.md)** collection returned by the **Selection** property to display the senders of all mail items selected in the active explorer window. To run this example, you need to have at least one mail item selected in the active Explorer window.


 **Note**  You might receive an error if you select items other than a mail item such as task request as the  **SenderName** property does not exist for a **TaskRequestItem** object.


```vb
Sub GetSelectedItems() 
 
 Dim myOlExp As Outlook.Explorer 
 
 Dim myOlSel As Outlook.Selection 
 
 Dim MsgTxt As String 
 
 Dim x As Integer 
 
 
 
 MsgTxt = "You have selected items from: " 
 
 Set myOlExp = Application.ActiveExplorer 
 
 Set myOlSel = myOlExp.Selection 
 
 For x = 1 To myOlSel.Count 
 
 MsgTxt = MsgTxt &; myOlSel.Item(x).SenderName &; ";" 
 
 Next x 
 
 MsgBox MsgTxt 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-outlook.md)

