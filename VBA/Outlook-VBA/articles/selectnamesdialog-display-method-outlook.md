---
title: SelectNamesDialog.Display Method (Outlook)
keywords: vbaol11.chm826
f1_keywords:
- vbaol11.chm826
ms.prod: outlook
api_name:
- Outlook.SelectNamesDialog.Display
ms.assetid: a689dfca-e4f7-f1c0-03a1-71e7d7e310b7
ms.date: 06/08/2017
---


# SelectNamesDialog.Display Method (Outlook)

Displays the  **Select Names** dialog box.


## Syntax

 _expression_ . **Display**

 _expression_ A variable that represents a **SelectNamesDialog** object.


### Return Value

A  **Boolean** value that is **True** if the user has clicked **OK**, and  **False** if the user has clicked **Cancel** or the Close icon.


## Remarks

When displaying the  **Select Names** dialog box, **Display** uses the previous location and size (indicated by the top, left, width, and height) of the dialog box.

The  **Select Names** dialog box is modal, meaning that code execution will halt until the user clicks **OK**,  **Cancel**, or the close icon.

You should detect for error conditions that include insufficient memory or another message or dialog box is open.


## Example

The following code sample shows how to create a mail item, allow the user to select recipients from the Exchange Global Address List in the  **Select Names** dialog box, and if the user has selected recipients that can be completely resolved, then send the mail item.


```vb
Sub SelectRecipients() 
 Dim oMsg As MailItem 
 Set oMsg = Application.CreateItem(olMailItem) 
 Dim oDialog As SelectNamesDialog 
 Set oDialog = Application.Session.GetSelectNamesDialog 
 With oDialog 
 .InitialAddressList = _ 
 Application.Session.GetGlobalAddressList 
 .Recipients = oMsg.Recipients 
 If .Display Then 
 'Recipients Resolved 
 oMsg.Subject = "Hello" 
 oMsg.Send 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[SelectNamesDialog Object](selectnamesdialog-object-outlook.md)

