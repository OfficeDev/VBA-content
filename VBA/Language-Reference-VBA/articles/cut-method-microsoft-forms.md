---
title: Cut Method (Microsoft Forms)
keywords: fm20.chm2000290
f1_keywords:
- fm20.chm2000290
ms.prod: office
ms.assetid: 9eea6f19-557d-2ae0-4e22-2f40b4d01caf
ms.date: 06/08/2017
---


# Cut Method (Microsoft Forms)



Removes selected information from an object and transfers it to the Clipboard.
 **Syntax**
 _object_. **Cut**
The  **Cut** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
 **Remarks**
For a  **ComboBox** or **TextBox**, the **Cut** method removes currently selected text in the control to the Clipboard. This method does not require that the control have the[focus](vbe-glossary.md).
On a  **Page**, **Frame**, or form, **Cut** removes currently selected controls to the Clipboard. This method only removes controls created at[run time](vbe-glossary.md).

