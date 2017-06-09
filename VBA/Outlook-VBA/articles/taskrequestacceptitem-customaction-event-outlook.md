---
title: TaskRequestAcceptItem.CustomAction Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.CustomAction
ms.assetid: 7a5d4fda-2f49-8263-d724-0e99867ed655
ms.date: 06/08/2017
---


# TaskRequestAcceptItem.CustomAction Event (Outlook)

Occurs when a custom action of an item (which is an instance of the parent object) executes.


## Syntax

 _expression_ . **CustomAction**( **_Action_** , **_Response_** , **_Cancel_** )

 _expression_ A variable that represents a **TaskRequestAcceptItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Action_|Required| **Object**|The  **[Action](action-object-outlook.md)** object.|
| _Response_|Required| **Object**|The newly created item resulting from the custom action.|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True** , the custom action is not completed.|

## Remarks

The  **Action** object and the newly created item resulting from the custom action are passed to the event.

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the custom action operation is not completed.


## See also


#### Concepts


[TaskRequestAcceptItem Object](taskrequestacceptitem-object-outlook.md)

