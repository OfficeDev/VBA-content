---
title: OptionButton Object (Outlook Forms Script)
keywords: olfm10.chm2000580
f1_keywords:
- olfm10.chm2000580
ms.prod: outlook
ms.assetid: 8009dd64-44b5-3b66-e8d4-e3535e014396
ms.date: 06/08/2017
---


# OptionButton Object (Outlook Forms Script)

Shows the selection status of one item in a group of choices.


## Remarks

Use an  **OptionButton** to show whether a single item in a group is selected. Note that each **OptionButton** in a **[Frame](frame-object-outlook-forms-script.md)** is mutually exclusive.

If an  **OptionButton** is bound to a data source, the **OptionButton** can show the value of that data source as either Yes/No, True/False, or On/Off. If the user selects the **OptionButton**, the current setting is Yes, True, or On. If the user does not select the  **OptionButton**, the setting is No, False, or Off. For example, an  **OptionButton** in an inventory-tracking application might show whether an item is discontinued. If the **OptionButton** is bound to a data source, then changing the setting changes the value of that data source. A disabled **OptionButton** is dimmed and does not show a value.

Depending on the value of the  **[TripleState](optionbutton-triplestate-property-outlook-forms-script.md)** property, an **OptionButton** can also have a **Null** value.

You can also use an  **OptionButton** inside a group box to select one or more of a group of related items. For example, you can create an order form with a list of available items, with an **OptionButton** preceding each item. The user can select a particular item by checking the corresponding **OptionButton** **OptionButton**.

The default property for an  **OptionButton** is the **[Value](optionbutton-value-property-outlook-forms-script.md)** property.


