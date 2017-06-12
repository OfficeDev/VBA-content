---
title: IntegralHeight Property
keywords: fm20.chm2001320
f1_keywords:
- fm20.chm2001320
ms.prod: office
api_name:
- Office.IntegralHeight
ms.assetid: 1aeec970-ef48-a9e8-f130-1ac51c61d026
ms.date: 06/08/2017
---


# IntegralHeight Property



Indicates whether a  **ListBox** or **TextBox** displays full lines of text in a list or partial lines.
 **Syntax**
 _object_. **IntegralHeight** [= _Boolean_ ]
The  **IntegralHeight** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Whether the list displays partial lines of text.|
 **Settings**
The settings for  _Boolean_ are:


|**Value**|**Description**|
|:-----|:-----|
|**True**|The list resizes itself to display only complete items (default).|
|**False**|The list does not resize itself even if the item is too tall to display completely.|
 **Remarks**
The  **IntegralHeight** property relates to the height of the list, just as the **AutoSize** property relates to the width of the list.
If  **IntegralHeight** is **True**, the list box automatically resizes when necessary to show full rows. If **False**, the list remains a fixed size; if items are taller than the available space in the list, the entire item is not shown.

