---
title: CanPaste Property
keywords: fm20.chm2000850
f1_keywords:
- fm20.chm2000850
ms.prod: office
api_name:
- Office.CanPaste
ms.assetid: 697a2f98-8c42-663c-9ff7-0330d3977c43
ms.date: 06/08/2017
---


# CanPaste Property



Specifies whether the Clipboard contains data that the object supports.
 **Syntax**
 _object_. **CanPaste**
The  **CanPaste** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
 **Return Values**
The  **CanPaste** property return values are:


|**Value**|**Description**|
|:-----|:-----|
|**True**|The object underneath the mouse pointer can receive information pasted from the Clipboard (default).|
|**False**|The object underneath the mouse pointer cannot receive information pasted from the Clipboard.|
 **Remarks**
 **CanPaste** is read-only.
If the Clipboard data is in a [format](glossary-vba.md) that the current[target](glossary-vba.md) object does not support, the **CanPaste** property is **False**. For example, if you try to paste a bitmap into an object that only supports text, **CanPaste** will be **False**.

