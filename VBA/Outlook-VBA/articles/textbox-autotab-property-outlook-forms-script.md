---
title: TextBox.AutoTab Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 4c7d917b-178b-04f2-9d9c-bf736eb9ad37
ms.date: 06/08/2017
---


# TextBox.AutoTab Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether an automatic tab occurs when a user enters the maximum allowable number of characters into a **[TextBox](textbox-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **AutoTab**

 _expression_A variable that represents a  **TextBox** object.


## Remarks

 **True** if tab occurs, **False** otherwise (default).

The  **[MaxLength](textbox-maxlength-property-outlook-forms-script.md)** property specifies the maximum number of characters allowed in a **TextBox**.

You can specify the  **AutoTab** property for a **TextBox** on a form for which you usually enter a set number of characters. Once a user enters the maximum number of characters, the focus automatically moves to the next control in the tab order. For example, if a **TextBox** displays inventory stock numbers that are always five characters long, you can use **MaxLength** to specify the maximum number of characters to enter into the **TextBox** and **AutoTab** to automatically tab to the next control after the user enters five characters.


