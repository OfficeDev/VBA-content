---
title: ToggleButton Object (Outlook Forms Script)
keywords: olfm10.chm2000680
f1_keywords:
- olfm10.chm2000680
ms.prod: outlook
ms.assetid: 01ce5640-9f19-3c0e-1aa4-96d87074bf8b
ms.date: 06/08/2017
---


# ToggleButton Object (Outlook Forms Script)

Shows the selection state of an item.


## Remarks

Use a  **ToggleButton** to show whether an item is selected. If a **ToggleButton** is bound to a data source, the **ToggleButton** shows the current value of that data source as either Yes/No, True/False, On/Off, or some other choice of two settings. If the user selects the **ToggleButton**, the current setting is Yes, True, or On. If the user does not select the  **ToggleButton**, the setting is No, False, or Off. If the  **ToggleButton** is bound to a data source, changing the setting changes the value of that data source. A disabled **ToggleButton** shows a value, but is dimmed and does not allow changes from the user interface.

You can also use a  **ToggleButton** inside a **[Frame](frame-object-outlook-forms-script.md)** to select one or more of a group of related items. For example, you can create an order form with a list of available items, with a **ToggleButton** preceding each item. The user can select a particular item by selecting the appropriate **ToggleButton**.

The default property of a  **ToggleButton** is the **[Value](togglebutton-value-property-outlook-forms-script.md)** property.

The only event for a  **ToggleButton** is the **[Click](togglebutton-click-event-outlook-forms-script.md)** event.


