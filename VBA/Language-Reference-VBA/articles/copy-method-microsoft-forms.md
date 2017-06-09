---
title: Copy Method (Microsoft Forms)
keywords: fm20.chm2000280
f1_keywords:
- fm20.chm2000280
ms.prod: office
ms.assetid: debe6ebb-e382-acd8-d8f3-c7808a47bccd
ms.date: 06/08/2017
---


# Copy Method (Microsoft Forms)



Copies the contents of an object to the Clipboard.
 **Syntax**
 _object_. **Copy**
The  **Copy** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
 **Remarks**
The original content remains on the object.
The actual content that is copied depends on the object. For example, on a  **Page**, the **Copy** method copies the currently selected control or controls. On a **TextBox** or **ComboBox**, it copies the currently selected text.
Using  **Copy** for a form, **Frame**, or **Page** copies the currently-active control.

