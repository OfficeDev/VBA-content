---
title: MailItem.Read Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MailItem.Read
ms.assetid: f20ec6d1-a2b4-9af3-66be-5398dc059c90
ms.date: 06/08/2017
---


# MailItem.Read Event (Outlook)

Occurs when an instance of the parent object is opened for editing by the user. 


## Syntax

 _expression_ . **Read**

 _expression_ A variable that represents a **MailItem** object.


## Remarks

The  **Read** event differs from the **[Open](mailitem-open-event-outlook.md)** event in that **Read** occurs whenever the user selects the item in a view that supports in-cell editing as well as when the item is being opened in an **[Inspector](inspector-object-outlook.md)** .


## Example

This Visual Basic for Applications (VBA) example uses the  **Read** event to increment a counter that tracks how often an item is read.


```vb
Public WithEvents myItem As Outlook.MailItem 
 
 
 
Sub Initialize_handler() 
 
 Set myItem = Application.ActiveExplorer.CurrentFolder.Items(1) 
 
 myItem.Display 
 
End Sub 
 
 
 
Sub myItem_Read() 
 
 Dim myProperty As Outlook.UserProperty 
 
 Set myProperty = myItem.UserProperties("ReadCount") 
 
 If (myProperty Is Nothing) Then 
 
 Set myProperty = myItem.UserProperties.Add("ReadCount", olNumber) 
 
 End If 
 
 myProperty.Value = myProperty.Value + 1 
 
 myItem.Save 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

