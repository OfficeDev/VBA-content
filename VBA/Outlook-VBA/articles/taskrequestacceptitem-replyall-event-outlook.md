---
title: TaskRequestAcceptItem.ReplyAll Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.ReplyAll
ms.assetid: 3bdca337-f106-b03f-c365-03d63aa22be8
ms.date: 06/08/2017
---


# TaskRequestAcceptItem.ReplyAll Event (Outlook)

Occurs when the user selects the  **ReplyAll** action for an item (which is an instance of the parent object).


## Syntax

 _expression_ . **ReplyAll**( **_Response_** , **_Cancel_** )

 _expression_ A variable that represents a **TaskRequestAcceptItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Response_|Required| **Object**|The new item being sent in response to the original message.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the reply all operation is not completed and the new item is not displayed.|

## Remarks

Returns the reply as a  **[MailItem](mailitem-object-outlook.md)** object.


## See also


#### Concepts


[TaskRequestAcceptItem Object](taskrequestacceptitem-object-outlook.md)

