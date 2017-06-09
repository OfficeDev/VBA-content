---
title: TextBox.TabKeyBehavior Property (Outlook Forms Script)
keywords: olfm10.chm2002020
f1_keywords:
- olfm10.chm2002020
ms.prod: outlook
ms.assetid: 5b8bdc3c-9000-a7fd-af39-743cc117e02d
ms.date: 06/08/2017
---


# TextBox.TabKeyBehavior Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether tabs are allowed in the edit region. Read/write.


## Syntax

 _expression_. **TabKeyBehavior**

 _expression_A variable that represents a  **TextBox** object.


## Remarks

 **True** if pressing **TAB** inserts a tab character in the edit region. **False** if pressing **TAB** moves the focus to the next object in the tab order (default).

The  **TabKeyBehavior** and **[MultiLine](textbox-multiline-property-outlook-forms-script.md)** properties are closely related. The values described above only apply if **MultiLine** is **True**. If  **MultiLine** is **False**, pressing  **TAB** always moves the focus to the next control in the tab order regardless of the value of **TabKeyBehavior**.

The effect of pressing  **CTRL+TAB** also depends on the value of **MultiLine**. If  **MultiLine** is **True**, pressing  **CTRL+TAB** creates a new line regardless of the value of **TabKeyBehavior**. If  **MultiLine** is **False**, pressing  **CTRL+TAB** has no effect.


