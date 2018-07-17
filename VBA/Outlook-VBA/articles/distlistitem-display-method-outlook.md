---
title: DistListItem.Display Method (Outlook)
keywords: vbaol11.chm1143
f1_keywords:
- vbaol11.chm1143
ms.prod: outlook
api_name:
- Outlook.DistListItem.Display
ms.assetid: 76a708ac-abfe-2a1c-85b3-624eb3b0514a
ms.date: 06/08/2017
---


# DistListItem.Display Method (Outlook)

Displays a new  **[Inspector](inspector-object-outlook.md)** object for the item.


## Syntax

 _expression_ . **Display**( **_Modal_** )

 _expression_ A variable that represents a **DistListItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Modal_|Optional| **Variant**| **True** to make the window modal. The default value is **False** .|

## Remarks

The  **Display** method is supported for explorer and inspector windows for the sake of backward compatibility. To activate an explorer or inspector window, use the **[Activate](inspector-activate-method-outlook.md)** method.

If you attempt to open an "unsafe" file system object (or "freedoc" file) by using the Microsoft Outlook object model, you receive the  **E_FAIL** return code in the C or C++ programming languages. In Outlook 2000 and earlier, you could open an "unsafe" file system object by using the **Display** method.


## See also


#### Concepts


[DistListItem Object](distlistitem-object-outlook.md)

