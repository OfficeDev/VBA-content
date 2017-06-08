---
title: SharingItem.Display Method (Outlook)
keywords: vbaol11.chm626
f1_keywords:
- vbaol11.chm626
ms.prod: outlook
api_name:
- Outlook.SharingItem.Display
ms.assetid: f243424e-06d7-8953-a19d-13f4f44dcabe
ms.date: 06/08/2017
---


# SharingItem.Display Method (Outlook)

Displays a new  **[Inspector](inspector-object-outlook.md)** object for the **[SharingItem](sharingitem-object-outlook.md)** .


## Syntax

 _expression_ . **Display**( **_Modal_** )

 _expression_ A variable that represents a **SharingItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Modal_|Optional| **Variant**| **True** to make the window modal. The default value is **False** .|

## Remarks

The  **Display** method is supported for explorer and inspector windows for the sake of backward compatibility. To activate an explorer or inspector window, use the **[Activate](inspector-activate-method-outlook.md)** method.

If you attempt to open an "unsafe" file system object (or "freedoc" file) by using the Microsoft Outlook object model, you receive the  **E_FAIL** return code in the C or C++ programming languages. In Outlook 2000 and earlier, you could open an "unsafe" file system object by using the **Display** method.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

