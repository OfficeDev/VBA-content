---
title: TaskRequestAcceptItem.Send Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.Send
ms.assetid: 47bbd12c-699e-8fb5-697d-e52e3acd12b5
ms.date: 06/08/2017
---


# TaskRequestAcceptItem.Send Event (Outlook)

Occurs when the user selects the  **Send** action for an item (which is an instance of the parent object).


## Syntax

 _expression_ . **Send**( **_Cancel_** )

 _expression_ A variable that represents a **TaskRequestAcceptItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True** , the send operation is not completed and the inspector is left open.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the item is not sent.


## Example

This Visual Basic for Applications (VBA) example uses the  **Send** event and sends an item with an automatic expiration date. The sample code must be placed in a class module such as `ThisOutlookSession`, and the  `SendMyMail` procedure must be called before the event procedure can be called by Microsoft Outlook.


```vb
Public WithEvents myItem As Outlook.MailItem 
 
 
 
Sub SendMyMail() 
 
 Set myItem = Outlook.CreateItem(olMailItem) 
 
 myItem.To = "Dan Wilson" 
 
 myItem.Subject = "Data files information" 
 
 myItem.Send 
 
End Sub 
 
 
 
Private Sub myItem_Send(Cancel As Boolean) 
 
 myItem.ExpiryTime = #2/2/2003 4:00:00 PM# 
 
End Sub
```


## See also


#### Concepts


[TaskRequestAcceptItem Object](taskrequestacceptitem-object-outlook.md)

