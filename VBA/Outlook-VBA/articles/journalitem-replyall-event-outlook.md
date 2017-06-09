---
title: JournalItem.ReplyAll Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.JournalItem.ReplyAll
ms.assetid: 86ab09f8-92f5-320e-9ec0-3be1f63c4583
ms.date: 06/08/2017
---


# JournalItem.ReplyAll Event (Outlook)

Occurs when the user selects the  **ReplyAll** action for an item, or when the **ReplyAll** method is called for the item, which is an instance of the parent object.


## Syntax

 _expression_ . **ReplyAll**( **_Response_** , **_Cancel_** )

 _expression_ A variable that represents a **JournalItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Response_|Required| **Object**|The new item being sent in response to the original message.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the reply all operation is not completed and the new item is not displayed.|

## Remarks

Returns the reply as a  **[MailItem](mailitem-object-outlook.md)** object.


## See also


#### Concepts


[JournalItem Object](journalitem-object-outlook.md)

