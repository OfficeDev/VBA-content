---
title: DocumentItem.Send Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.DocumentItem.Send
ms.assetid: 7e77c1c3-f6dd-13d1-ed76-b37e7dd6e82a
ms.date: 06/08/2017
---


# DocumentItem.Send Event (Outlook)

Occurs when the user selects the  **Send** action for an item (which is an instance of the parent object).


## Syntax

 _expression_ . **Send**( **_Cancel_** )

 _expression_ A variable that represents a **DocumentItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True** , the send operation is not completed and the inspector is left open.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the item is not sent.


## See also


#### Concepts


[DocumentItem Object](documentitem-object-outlook.md)

