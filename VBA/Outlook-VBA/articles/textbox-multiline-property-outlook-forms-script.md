---
title: TextBox.MultiLine Property (Outlook Forms Script)
keywords: olfm10.chm2001560
f1_keywords:
- olfm10.chm2001560
ms.prod: outlook
ms.assetid: f42aadc5-ecd9-090b-cdf0-aba0a1a024b2
ms.date: 06/08/2017
---


# TextBox.MultiLine Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether a control can accept and display multiple lines of text. Read/write.


## Syntax

 _expression_. **MultiLine**

 _expression_A variable that represents a  **TextBox** object.


## Remarks

True if the text is displayed across multiple lines (default). Falase if the text is not displayed across multiple lines.

A multiline  **[TextBox](textbox-object-outlook-forms-script.md)** allows absolute line breaks and adjusts its quantity of lines to accommodate the amount of text it holds. If needed, a multiline control can have vertical scroll bars.

A single-line  **TextBox** doesn't allow absolute line breaks and doesn't use vertical scroll bars.

For controls that support the  **MultiLine** property as well as the **[WordWrap](textbox-wordwrap-property-outlook-forms-script.md)** property, **WordWrap** is ignored when **MultiLine** is **False**.

Single-line controls ignore the value of the  **WordWrap** property.

If you change  **MultiLine** to **False** in a multiline **TextBox**, all the characters in the  **TextBox** will be combined into one line, including non-printing characters (such as carriage returns and new-lines).

The  **[EnterKeyBehavior](textbox-enterkeybehavior-property-outlook-forms-script.md)** and **MultiLine** properties are closely related. The **EnterKeyBehavior** values of **True** and **False** only apply if **MultiLine** is **True**. If  **MultiLine** is **False**, pressing  **ENTER** always moves the focus to the next control in the tab order regardless of the value of **EnterKeyBehavior**.

The effect of pressing  **CTRL+ENTER** also depends on the value of **MultiLine**. If  **MultiLine** is **True**, pressing  **CTRL+ENTER** creates a new line regardless of the value of **EnterKeyBehavior**. If  **MultiLine** is **False**, pressing  **CTRL+ENTER** has no effect.

The  **[TabKeyBehavior](textbox-tabkeybehavior-property-outlook-forms-script.md)** and **MultiLine** properties are closely related. The values described above only apply if **MultiLine** is **True**. If  **MultiLine** is **False**, pressing  **TAB** always moves the focus to the next control in the tab order regardless of the value of **TabKeyBehavior**.

The effect of pressing  **CTRL+TAB** also depends on the value of **MultiLine**. If  **MultiLine** is **True**, pressing  **CTRL+TAB** creates a new line regardless of the value of **TabKeyBehavior**. If  **MultiLine** is **False**, pressing  **CTRL+TAB** has no effect.


