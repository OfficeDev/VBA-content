---
title: Tips on using text boxes
keywords: fm20.chm5225199
f1_keywords:
- fm20.chm5225199
ms.prod: office
ms.assetid: 4f2c565a-50ba-0295-d8bf-92d316ea25af
ms.date: 06/08/2017
---


# Tips on using text boxes

The  **TextBox** is a flexible control governed by the following properties: **Text**, **MultiLine**, **WordWrap**, and **AutoSize**.

 **Text** contains the text that's displayed in the text box.

 **MultiLine** controls whether the **TextBox** can display text as a single line or as multiple lines. Newline characters identify where one line ends and another begins. If **MultiLine** is **False**, then the text is truncated instead of wrapped.

 **WordWrap** allows the **TextBox** to wrap lines of text that are longer than the width of the **TextBox** into shorter lines that fit.
If you do not use  **WordWrap**, the **TextBox** starts a new line of text when it encounters a newline character in the text. If **WordWrap** is turned off, you can have text lines that do not fit completely in the **TextBox**. The **TextBox** displays the portions of text that fit inside its width and truncates the portions of text that do not fit. **WordWrap** is not applicable unless **MultiLine** is **True**.
 **AutoSize** controls whether the **TextBox** adjusts to display all of the text. When using **AutoSize** with a **TextBox**, the width of the **TextBox** shrinks or expands according to the amount of text in the **TextBox** and the font size used to display the text.
 **AutoSize** works well in the following situations:


- Displaying a caption of one or more lines.
    
- Displaying the contents of a single-line  **TextBox**.
    
- Displaying the contents of a multiline  **TextBox** that is read-only to the user.
    


 **Note**  Avoid using  **AutoSize** with an empty **TextBox** that also uses the **MultiLine** and **WordWrap** properties. When the user enters text into a **TextBox** with these properties, the **TextBox** automatically sizes to a long narrow box one character wide and as long as the line of text.


