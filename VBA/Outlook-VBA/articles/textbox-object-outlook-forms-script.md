---
title: TextBox Object (Outlook Forms Script)
keywords: olfm10.chm2000670
f1_keywords:
- olfm10.chm2000670
ms.prod: outlook
ms.assetid: 4a0e4a3d-beca-9f94-7e27-469c4bafe250
ms.date: 06/08/2017
---


# TextBox Object (Outlook Forms Script)

Displays information from a user or from an organized set of data.


## Remarks

A  **TextBox** is the control most commonly used to display information entered by a user. Also, it can display a set of data, such as a table, query, worksheet, or a calculation result. If a **TextBox** is bound to a data source, then changing the contents of the **TextBox** also changes the value of the bound data source.

Formatting applied to any piece of text in a  **TextBox** will affect all text in the control. For example, if you change the font or point size of any character in the control, the change will affect all characters in the control.

The default property for a  **TextBox** is the **[Value](textbox-value-property-outlook-forms-script.md)** property.


### Tips on using text boxes

The  **TextBox** is a flexible control governed by the following properties: **[Text](textbox-text-property-outlook-forms-script.md)**,  **[MultiLine](textbox-multiline-property-outlook-forms-script.md)**,  **[WordWrap](textbox-wordwrap-property-outlook-forms-script.md)**, and  **[AutoSize](textbox-autosize-property-outlook-forms-script.md)**.

 **Text** contains the text that's displayed in the text box.

 **MultiLine** controls whether the **TextBox** can display text as a single line or as multiple lines. Newline characters identify where one line ends and another begins. If **MultiLine** is **False** (the default value), the text is truncated instead of wrapped.

 **WordWrap** allows the **TextBox** to wrap lines of text that are longer than the width of the **TextBox** into shorter lines that fit. The default value is **True**.

If you do not use  **WordWrap**, the  **TextBox** starts a new line of text when it encounters a newline character in the text. If **WordWrap** is turned off, you can have text lines that do not fit completely in the **TextBox**. The  **TextBox** displays the portions of text that fit inside its width and truncates the portions of text that do not fit. **WordWrap** is not applicable unless **MultiLine** is **True**.

 **AutoSize** controls whether the **TextBox** adjusts to display all of the text. When using **AutoSize** with a **TextBox**, the width of the  **TextBox** shrinks or expands according to the amount of text in the **TextBox** and the font size used to display the text. The default value is **False**.

 **AutoSize** works well in the following situations:


- Displaying a caption of one or more lines.
    
- Displaying the contents of a single-line  **TextBox**.
    
- Displaying the contents of a multiline  **TextBox** that is read-only to the user.
    
Avoid using  **AutoSize** with an empty **TextBox** that also uses the **MultiLine** and **WordWrap** properties. When the user enters text into a **TextBox** with these properties, the **TextBox** automatically sizes to a long narrow box one character wide and as long as the line of text.


