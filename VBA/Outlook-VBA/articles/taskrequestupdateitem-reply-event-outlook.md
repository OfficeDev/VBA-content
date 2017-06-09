---
title: TaskRequestUpdateItem.Reply Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem.Reply
ms.assetid: b6c07e2a-04a7-bd0a-cb09-9b4ddcbf97ae
ms.date: 06/08/2017
---


# TaskRequestUpdateItem.Reply Event (Outlook)

Occurs when the user selects the  **Reply** action for an item (which is an instance of the parent object).


## Syntax

 _expression_ . **Reply**( **_Response_** , **_Cancel_** )

 _expression_ A variable that represents a **TaskRequestUpdateItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Response_|Required| **Object**|The new item being sent in response to the original message.|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True** , the reply operation is not completed and the new item is not displayed.|

## Remarks

Returns the reply as a  **[MailItem](mailitem-object-outlook.md)** object.

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the reply action is not completed and the new item is not displayed.


## See also


#### Concepts


[TaskRequestUpdateItem Object](taskrequestupdateitem-object-outlook.md)

