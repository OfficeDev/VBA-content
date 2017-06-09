---
title: SharingItem.Forward Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.SharingItem.Forward
ms.assetid: b9f8cb45-e4e8-2eb5-c892-9d718bffae74
ms.date: 06/08/2017
---


# SharingItem.Forward Event (Outlook)

Occurs when the user selects the  **Forward** action for an item, or when the **[Forward](sharingitem-forward-method-outlook.md)** method is called for the item, which is an instance of the parent object.


## Syntax

 _expression_ . **Forward**( **_Forward_** , **_Cancel_** )

 _expression_ An expression that returns a **SharingItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Forward_|Required| **Object**|The new item being forwarded.|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True** , the forward operation is not completed and the new item is not displayed.|

## Remarks

In VBScript, if you set the return value of this function to  **False** , the forward action is not completed and the new item is not displayed.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

