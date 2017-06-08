---
title: ListBox Object (Outlook Forms Script)
keywords: olfm10.chm2000560
f1_keywords:
- olfm10.chm2000560
ms.prod: outlook
ms.assetid: f56ba480-f8fe-6d12-265e-3b0a9838af97
ms.date: 06/08/2017
---


# ListBox Object (Outlook Forms Script)

Displays a list of values and lets you select one or more.


## Remarks

If the  **ListBox** is bound to a data source, the **ListBox** stores the selected value in that data source.

The  **ListBox** can either appear as a list or as a group of **[OptionButton](optionbutton-object-outlook-forms-script.md)** controls or **[CheckBox](checkbox-object-outlook-forms-script.md)** controls.

The default property for a  **ListBox** is the **[Value](listbox-value-property-outlook-forms-script.md)** property.

The default event for a  **ListBox** is the **[Click](listbox-click-event-outlook-forms-script.md)** event.

You can't drop text into a drop-down  **ListBox**.


### ListBox styles

You can choose between two presentation styles for a  **ListBox**. This is expressed by the  **[ListStyle](listbox-liststyle-property-outlook-forms-script.md)** property. Each style provides different ways for users to select items in the list.

If the style is 0, each item is on a separate row; the user selects an item by highlighting one or more rows.

If the style is 1, an  **OptionButton** or **CheckBox** appears at the beginning of each row. With this style, the user selects an item by clicking the option button or check box. Check boxes appear only when the **[MultiSelect](listbox-multiselect-property-outlook-forms-script.md)** property is **True**.


