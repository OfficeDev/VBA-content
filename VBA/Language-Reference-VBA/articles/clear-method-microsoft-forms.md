---
title: Clear Method (Microsoft Forms)
keywords: fm20.chm5224955
f1_keywords:
- fm20.chm5224955
ms.prod: office
ms.assetid: c0fe2f8c-1af1-6977-e794-38f9fa40deac
ms.date: 06/08/2017
---


# Clear Method (Microsoft Forms)



Removes all objects from an object or [collection](vbe-glossary.md).
 **Syntax**
 _object_. **Clear**
The  **Clear** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
 **Remarks**
For a  **MultiPage** or **TabStrip**, the **Clear** method deletes individual pages or tabs.
For a  **ListBox** or **ComboBox**, **Clear** removes all entries in the list.
For a  **Controls** collection, **Clear** deletes controls that were created at[run time](vbe-glossary.md) with the **Add** method. Using **Clear** on controls created at[design time](vbe-glossary.md) causes an error.
If the control is bound to data, the  **Clear** method fails.

