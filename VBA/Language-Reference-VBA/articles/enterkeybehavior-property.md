---
title: EnterKeyBehavior Property
keywords: fm20.chm5225037
f1_keywords:
- fm20.chm5225037
ms.prod: office
api_name:
- Office.EnterKeyBehavior
ms.assetid: 720a6b10-f021-e623-7f63-f52081bcafd1
ms.date: 06/08/2017
---


# EnterKeyBehavior Property



Defines the effect of pressing ENTER in a  **TextBox**.
 **Syntax**
 _object_. **EnterKeyBehavior** [= _Boolean_ ]
The  **EnterKeyBehavior** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Specifies the effect of pressing ENTER.|
 **Settings**
The settings for  _Boolean_ are:


|**Value**|**Description**|
|:-----|:-----|
|**True**|Pressing ENTER creates a new line.|
|**False**|Pressing ENTER moves the focus to the next object in the tab order (default).|
 **Remarks**
The  **EnterKeyBehavior** and **MultiLine** properties are closely related. The values described above only apply if **MultiLine** is **True**. If **MultiLine** is **False**, pressing ENTER always moves the[focus](vbe-glossary.md) to the next control in the[tab order](vbe-glossary.md) regardless of the value of **EnterKeyBehavior**.
The effect of pressing CTRL+ENTER also depends on the value of  **MultiLine**. If **MultiLine** is **True**, pressing CTRL+ENTER creates a new line regardless of the value of **EnterKeyBehavior**. If **MultiLine** is **False**, pressing CTRL+ENTER has no effect.

