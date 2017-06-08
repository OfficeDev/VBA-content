---
title: MeetingItem.Send Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MeetingItem.Send
ms.assetid: 9dc87c39-d209-dc06-86e8-ce00f9cb152f
ms.date: 06/08/2017
---


# MeetingItem.Send Event (Outlook)

Occurs when the user selects the  **Send** action for an item, or when the **Send** method is called for the item, which is an instance of the parent object.


## Syntax

 _expression_ . **Send**( **_Cancel_** )

 _expression_ A variable that represents a **MeetingItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True** , the send operation is not completed and the inspector is left open.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the item is not sent.


## See also


#### Concepts


[MeetingItem Object](meetingitem-object-outlook.md)

