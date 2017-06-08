---
title: CheckBox Object (Outlook Forms Script)
keywords: olfm10.chm2000470
f1_keywords:
- olfm10.chm2000470
ms.prod: outlook
ms.assetid: 1834855b-f96c-aaa1-24ce-81d1e4e4e1db
ms.date: 06/08/2017
---


# CheckBox Object (Outlook Forms Script)

Displays the selection state of an item.


## Remarks

Use a  **CheckBox** to give the user a choice between two values such as **Yes/No**,  **True/False**, or  **On/Off**. When the user selects a  **CheckBox**, it displays a special mark (such as an  **X**) and its current setting is  **Yes**,  **True**, or  **On**. If the user does not select the  **CheckBox**, it is empty and its setting is  **No**,  **False**, or Off. Depending on the value of the  **[TripleState](checkbox-triplestate-property-outlook-forms-script.md)** property, a **CheckBox** can also have a **Null** value.

If a  **CheckBox** is bound to a data source, changing the setting changes the value of that source. A disabled **CheckBox** shows the current value, but is dimmed and does not allow changes to the value from the user interface.

You can also use check boxes inside a group box to select one or more of a group of related items. For example, you can create an order form that contains a list of available items, with a  **CheckBox** preceding each item. The user can select a particular item or items by checking the corresponding **CheckBox**.

The default property of a  **CheckBox** is the **[Value](checkbox-value-property-outlook-forms-script.md)** property.

The  **[ListBox](listbox-object-outlook-forms-script.md)** also lets you put a check mark by selected options. Depending on your application, you can use the **ListBox** instead of using a group of **CheckBox** controls.


