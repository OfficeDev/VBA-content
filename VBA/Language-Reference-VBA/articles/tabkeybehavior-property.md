---
title: TabKeyBehavior Property
keywords: fm20.chm5225099
f1_keywords:
- fm20.chm5225099
ms.prod: office
api_name:
- Office.TabKeyBehavior
ms.assetid: 9019c946-8590-2538-fbf0-c9d131a78963
ms.date: 06/08/2017
---


# TabKeyBehavior Property



Determines whether tabs are allowed in the edit region.
 **Syntax**
 _object_. **TabKeyBehavior** [= _Boolean_ ]
The  **TabKeyBehavior** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. The effect of pressing TAB.|
 **Settings**
The settings for  _Boolean_ are:


|**Value**|**Description**|
|:-----|:-----|
|**True**|Pressing TAB inserts a tab character in the edit region.|
|**False**|Pressing TAB moves the focus to the next object in the tab order (default).|
 **Remarks**
The  **TabKeyBehavior** and **MultiLine** properties are closely related. The values described above only apply if **MultiLine** is **True**. If **MultiLine** is **False**, pressing TAB always moves the[focus](vbe-glossary.md) to the next control in the[tab order](vbe-glossary.md) regardless of the value of **TabKeyBehavior**.
The effect of pressing CTRL+TAB also depends on the value of  **MultiLine**. If **MultiLine** is **True**, pressing CTRL+TAB creates a new line regardless of the value of **TabKeyBehavior**. If **MultiLine** is **False**, pressing CTRL+TAB has no effect.

