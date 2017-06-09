---
title: Explorer.ActiveInlineResponse Property (Outlook)
keywords: vbaol11.chm3595
f1_keywords:
- vbaol11.chm3595
ms.assetid: fc38314d-7cff-44f4-9151-6129f918a721
ms.date: 06/08/2017
ms.prod: outlook
---


# Explorer.ActiveInlineResponse Property (Outlook)
Returns an item object representing the active inline response item in the explorer reading pane. Read-only.

## Syntax

 _expression_ . **ActiveInlineResponse**

 _expression_ A variable that represents an **[Explorer](explorer-object-outlook.md)** object.


## Remarks

You can use the same properties and methods of the [MailItem](mailitem-object-outlook.md) object on this item, except for the following:


- [MailItem.Actions](mailitem-actions-property-outlook.md) property
    
- [MailItem.Close](mailitem-close-method-outlook.md) method
    
- [MailItem.Copy](mailitem-copy-method-outlook.md) method
    
- [MailItem.Delete](mailitem-delete-method-outlook.md) method
    
- [MailItem.Forward](mailitem-forward-method-outlook.md) method
    
- [MailItem.Move](mailitem-move-method-outlook.md) method
    
- [MailItem.Reply](mailitem-reply-method-outlook.md) method
    
- [MailItem.ReplyAll](mailitem-replyall-method-outlook.md) method
    
- [MailItem.Send](mailitem-send-method-outlook.md) method
    
This property returns  **Null** ( **Nothing** in Visual Basic) if no inline response is visible in the Reading Pane.


## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

