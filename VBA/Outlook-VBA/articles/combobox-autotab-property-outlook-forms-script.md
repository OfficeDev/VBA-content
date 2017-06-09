---
title: ComboBox.AutoTab Property (Outlook Forms Script)
keywords: olfm10.chm2000750
f1_keywords:
- olfm10.chm2000750
ms.prod: outlook
ms.assetid: e6dc50c5-8766-21c5-3b4f-bd0b92882128
ms.date: 06/08/2017
---


# ComboBox.AutoTab Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether an automatic tab occurs when a user enters the maximum allowable number of characters into the text box portion of a **[ComboBox](combobox-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **AutoTab**

 _expression_A variable that represents a  **ComboBox** object.


## Remarks

 **True** if tab occurs, **False** otherwise (default).

The  **[MaxLength](combobox-maxlength-property-outlook-forms-script.md)** property specifies the maximum number of characters allowed in the text box portion of a **ComboBox**.

You can specify the  **AutoTab** property for a **ComboBox** on a form for which you usually enter a set number of characters. Once a user enters the maximum number of characters, the focus automatically moves to the next control in the tab order. For example, if a **ComboBox** displays inventory stock numbers that are always five characters long, you can use **MaxLength** to specify the maximum number of characters to enter into the **ComboBox** and **AutoTab** to automatically tab to the next control after the user enters five characters.


