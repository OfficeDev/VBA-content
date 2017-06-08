---
title: SharingItem.ReplyAll Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.SharingItem.ReplyAll
ms.assetid: 147f7da9-fa4b-b678-f600-25a8c6b540ec
ms.date: 06/08/2017
---


# SharingItem.ReplyAll Event (Outlook)

Occurs when the user selects the  **ReplyAll** action for an item, or when the **[ReplyAll](sharingitem-replyall-method-outlook.md)** method is called for the item, which is an instance of the parent object.


## Syntax

 _expression_ . **ReplyAll**( **_Response_** , **_Cancel_** )

 _expression_ An expression that returns a **SharingItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Response_|Required| **Object**|The new item being sent in response to the original message.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the reply all operation is not completed and the new item is not displayed.|

## Remarks

Returns the reply as a  **[MailItem](mailitem-object-outlook.md)** object.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

