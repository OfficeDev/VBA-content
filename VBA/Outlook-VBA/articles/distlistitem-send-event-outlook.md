---
title: DistListItem.Send Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.DistListItem.Send
ms.assetid: 8f92ff6e-9922-1460-0c9d-eba77dadbba1
ms.date: 06/08/2017
---


# DistListItem.Send Event (Outlook)

Occurs when the user selects the  **Send** action for an item (which is an instance of the parent object).


## Syntax

 _expression_ . **Send**( **_Cancel_** )

 _expression_ A variable that represents a **DistListItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True** , the send operation is not completed and the inspector is left open.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the item is not sent.


## See also


#### Concepts


[DistListItem Object](distlistitem-object-outlook.md)

