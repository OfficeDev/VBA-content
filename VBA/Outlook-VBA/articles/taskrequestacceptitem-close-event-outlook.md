---
title: TaskRequestAcceptItem.Close Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.Close
ms.assetid: 93504a80-71ae-ff95-1b39-5aef22dabdb8
ms.date: 06/08/2017
---


# TaskRequestAcceptItem.Close Event (Outlook)

Occurs when the inspector associated with an item (which is an instance of the parent object) is being closed.


## Syntax

 _expression_ . **Close**( **_Cancel_** )

 _expression_ A variable that represents a **TaskRequestAcceptItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True** , the close operation is not completed and the inspector is left open.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the close operation isn't completed and the inspector is left open.

If you use the  **[Close](taskrequestacceptitem-close-method-outlook.md)** method to fire this event, it can only be canceled if the **Close** method uses the **olPromptForSave** argument.


## See also


#### Concepts


[TaskRequestAcceptItem Object](taskrequestacceptitem-object-outlook.md)

