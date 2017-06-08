---
title: RemoteItem.Forward Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.RemoteItem.Forward
ms.assetid: f4af05e8-c0ea-915e-f49c-2470620e0735
ms.date: 06/08/2017
---


# RemoteItem.Forward Event (Outlook)

Occurs when the user selects the  **Forward** action for an item (which is an instance of the parent object).


## Syntax

 _expression_ . **Forward**( **_Forward_** , **_Cancel_** )

 _expression_ A variable that represents a **RemoteItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Forward_|Required| **Object**|The new item being forwarded.|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True** , the forward operation is not completed and the new item is not displayed.|

## Remarks

In VBScript, if you set the return value of this function to  **False** , the forward action is not completed and the new item is not displayed.


## See also


#### Concepts


[RemoteItem Object](remoteitem-object-outlook.md)

