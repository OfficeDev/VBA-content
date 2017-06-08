---
title: TaskRequestItem.ReplyAll Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.ReplyAll
ms.assetid: f68693c4-9cfd-2805-3559-4117c46f05c8
ms.date: 06/08/2017
---


# TaskRequestItem.ReplyAll Event (Outlook)

Occurs when the user selects the  **ReplyAll** action for an item (which is an instance of the parent object).


## Syntax

 _expression_ . **ReplyAll**( **_Response_** , **_Cancel_** )

 _expression_ A variable that represents a **TaskRequestItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Response_|Required| **Object**|The new item being sent in response to the original message.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the reply all operation is not completed and the new item is not displayed.|

## Remarks

Returns the reply as a  **[MailItem](mailitem-object-outlook.md)** object.


## See also


#### Concepts


[TaskRequestItem Object](taskrequestitem-object-outlook.md)

