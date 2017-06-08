---
title: OptionButton Control
keywords: fm20.chm5224984
f1_keywords:
- fm20.chm5224984
ms.prod: office
ms.assetid: 39ce3eb0-ecf1-4f1e-dbcb-a66d7d341615
ms.date: 06/08/2017
---


# OptionButton Control



Shows the selection status of one item in a [group](glossary-vba.md) of choices.
 **Remarks**
Use an  **OptionButton** to show whether a single item in a group is selected. Note that each **OptionButton** in a **Frame** is mutually exclusive.
If an  **OptionButton** is[bound](glossary-vba.md) to a[data source](glossary-vba.md), the  **OptionButton** can show the value of that data source as either _Yes_ / _No_, _True_ / _False_, or _On_ / _Off_. If the user selects the **OptionButton**, the current setting is _Yes_, _True_, or _On_; if the user does not select the **OptionButton**, the setting is _No_, _False_, or _Off_. For example, an **OptionButton** in an inventory-tracking application might show whether an item is discontinued. If the **OptionButton** is bound to a data source, then changing the settings changes the value of that data source. A disabled **OptionButton** is dimmed and does not show a value.
Depending on the value of the  **TripleState** property, an **OptionButton** can also have a[null](vbe-glossary.md) value.
You can also use an  **OptionButton** inside a group box to select one or more of a group of related items. For example, you can create an order form with a list of available items, with an **OptionButton** preceding each item. The user can select a particular item by checking the corresponding **OptionButton**.
The default property for an  **OptionButton** is the **Value** property.
The default event for an  **OptionButton** is the Click event.

## Related Topics

[ OptionButton Object](http://msdn.microsoft.com/library/5cff61be-6357-4db4-b381-b168626d5f28%28Office.15%29.aspx)


