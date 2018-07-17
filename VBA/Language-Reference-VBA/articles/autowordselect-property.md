---
title: AutoWordSelect Property
keywords: fm20.chm2000760
f1_keywords:
- fm20.chm2000760
ms.prod: office
api_name:
- Office.AutoWordSelect
ms.assetid: 24e9e8ff-5988-9ed3-4a2c-f3faa99248f9
ms.date: 06/08/2017
---


# AutoWordSelect Property



Specifies whether a word or a character is the basic unit used to extend a selection.
 **Syntax**
 _object_. **AutoWordSelect** [= _Boolean_ ]
The  **AutoWordSelect** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Specifies the basic unit used to extend a selection.|
 **Settings**
The settings for  _Boolean_ are:


|**Value**|**Description**|
|:-----|:-----|
|**True**|Uses a word as the basic unit (default).|
|**False**|Uses a character as the basic unit.|
 **Remarks**
The  **AutoWordSelect** property specifies how the selection extends or contracts in the edit region of a **TextBox** or **ComboBox**.
If the user places the insertion point in the middle of a word and then extends the selection while  **AutoWordSelect** is **True**, the selection includes the entire word.

