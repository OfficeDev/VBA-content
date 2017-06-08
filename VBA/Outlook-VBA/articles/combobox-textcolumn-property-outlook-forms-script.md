---
title: ComboBox.TextColumn Property (Outlook Forms Script)
keywords: olfm10.chm2002090
f1_keywords:
- olfm10.chm2002090
ms.prod: outlook
ms.assetid: 5ebf37ef-4cec-ec42-d42f-ab886b86e913
ms.date: 06/08/2017
---


# ComboBox.TextColumn Property (Outlook Forms Script)

Returns or sets a  **Variant** that identifies the column in a **[ComboBox](combobox-object-outlook-forms-script.md)** to display to the user. Read/write.


## Syntax

 _expression_. **TextColumn**

 _expression_A variable that represents a  **ComboBox** object.


## Remarks

Values for the  **TextColumn** property range from -1 to the number of columns in the list. The **TextColumn** value for the first column is 1, the value of the second column is 2, and so on. Setting **TextColumn** to 0 displays the **[ListIndex](combobox-listindex-property-outlook-forms-script.md)** values. Setting **TextColumn** to -1 displays the first column that has a **[ColumnWidths](combobox-columnwidths-property-outlook-forms-script.md)** value greater than 0.

When the user selects a row from a  **ComboBox**, the column referenced by  **TextColumn** is stored in the **[Text](combobox-text-property-outlook-forms-script.md)** property. For example, you could set up a multicolumn **ComboBox** that contains the names of holidays in one column and dates for the holidays in a second column. To present the holiday names to users, specify the first column as the **TextColumn**. To store the dates of the holidays, specify the second column as the  **[BoundColumn](combobox-boundcolumn-property-outlook-forms-script.md)**.

When the  **Text** property of a **ComboBox** changes (such as when a user types an entry into the control), the new text is compared to the column of data specified by **TextColumn**.


