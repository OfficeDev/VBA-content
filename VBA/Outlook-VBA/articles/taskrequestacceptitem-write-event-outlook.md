---
title: TaskRequestAcceptItem.Write Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.Write
ms.assetid: 005b0f33-1848-101b-2119-cb15eb51f411
ms.date: 06/08/2017
---


# TaskRequestAcceptItem.Write Event (Outlook)

Occurs when an instance of the parent object is saved, either explicitly (for example, using the  **[Save](taskrequestacceptitem-save-method-outlook.md)** or **[SaveAs](taskrequestacceptitem-saveas-method-outlook.md)** methods) or implicitly (for example, in response to a prompt when closing the item's inspector).


## Syntax

 _expression_ . **Write**( **_Cancel_** )

 _expression_ A variable that represents a **TaskRequestAcceptItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| (Not used in VBScript). **False** when the event occurs. If the event procedure sets this argument to **True** , the save operation is not completed.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the save operation is not completed.


## See also


#### Concepts


[TaskRequestAcceptItem Object](taskrequestacceptitem-object-outlook.md)

