---
title: MailItem.Reply Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MailItem.Reply
ms.assetid: 0bf6a21a-f667-9851-aeb0-dd6b9b83876e
ms.date: 06/08/2017
---


# MailItem.Reply Event (Outlook)

Occurs when the user selects the  **Reply** action for an item, or when the **Reply** method is called for the item, which is an instance of the parent object.


## Syntax

 _expression_ . **Reply**( **_Response_** , **_Cancel_** )

 _expression_ A variable that represents a **MailItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Response_|Required| **Object**|The new item being sent in response to the original message.|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True** , the reply operation is not completed and the new item is not displayed.|

## Remarks

Returns the reply as a  **[MailItem](mailitem-object-outlook.md)** object.

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the reply action is not completed and the new item is not displayed.


## Example

This Visual Basic for Applications (VBA) example uses the  **Reply** event and sets the **Sent Items** folder for the reply item to the folder in which the original item resides. To use this example, open an existing mailitem, run the `Initialize Handler()` procedure, then reply to the open item.


```vb
Public WithEvents myItem As MailItem 
 
 
 
Sub Initialize_Handler() 
 
 Set myItem = Application.ActiveInspector.CurrentItem 
 
End Sub 
 
 
 
Private Sub myItem_Reply(ByVal Response As Object, Cancel As Boolean) 
 
 Set Response.SaveSentMessageFolder = myItem.Parent 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

