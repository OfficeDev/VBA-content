---
title: TaskItem.Send Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskItem.Send
ms.assetid: f634105e-5351-6941-e915-ec63cd703b67
ms.date: 06/08/2017
---


# TaskItem.Send Event (Outlook)

Occurs when the user selects the  **Send** action for an item, or when the **Send** method is called for the item, which is an instance of the parent object.


## Syntax

 _expression_ . **Send**( **_Cancel_** )

 _expression_ A variable that represents a **TaskItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True** , the send operation is not completed and the inspector is left open.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the item is not sent.


## See also


#### Concepts


[TaskItem Object](taskitem-object-outlook.md)

