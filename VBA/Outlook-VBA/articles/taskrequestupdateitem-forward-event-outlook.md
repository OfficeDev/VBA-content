---
title: TaskRequestUpdateItem.Forward Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem.Forward
ms.assetid: c992a365-b36b-278d-5c93-32fa4b0f4993
ms.date: 06/08/2017
---


# TaskRequestUpdateItem.Forward Event (Outlook)

Occurs when the user selects the  **Forward** action for an item (which is an instance of the parent object).


## Syntax

 _expression_ . **Forward**( **_Forward_** , **_Cancel_** )

 _expression_ A variable that represents a **TaskRequestUpdateItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Forward_|Required| **Object**|The new item being forwarded.|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True** , the forward operation is not completed and the new item is not displayed.|

## Remarks

In VBScript, if you set the return value of this function to  **False** , the forward action is not completed and the new item is not displayed.


## See also


#### Concepts


[TaskRequestUpdateItem Object](taskrequestupdateitem-object-outlook.md)

