---
title: MeetingItem.Write Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MeetingItem.Write
ms.assetid: 22a52e41-cbc5-ced7-a942-ae06035aebbb
ms.date: 06/08/2017
---


# MeetingItem.Write Event (Outlook)

Occurs when an instance of the parent object is saved, either explicitly (for example, using the  **[Save](meetingitem-save-method-outlook.md)** or **[SaveAs](meetingitem-saveas-method-outlook.md)** methods) or implicitly (for example, in response to a prompt when closing the item's inspector).


## Syntax

 _expression_ . **Write**( **_Cancel_** )

 _expression_ A variable that represents a **MeetingItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| (Not used in VBScript). **False** when the event occurs. If the event procedure sets this argument to **True** , the save operation is not completed.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the save operation is not completed.


## See also


#### Concepts


[MeetingItem Object](meetingitem-object-outlook.md)

