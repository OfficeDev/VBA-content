---
title: CheckBox Control
keywords: fm20.chm5224977
f1_keywords:
- fm20.chm5224977
ms.prod: office
ms.assetid: 24d90604-51ec-7f7d-e679-52391b2c27c0
ms.date: 06/08/2017
---


# CheckBox Control



Displays the selection state of an item.
 **Remarks**
Use a  **CheckBox** to give the user a choice between two values such as _Yes_ / _No_, _True_ / _False_, or _On_ / _Off_. When the user selects a **CheckBox**, it displays a special mark (such as an X) and its current setting is _Yes_, _True_, or _On_; if the user does not select the **CheckBox**, it is empty and its setting is _No_, _False_, or _Off_. Depending on the value of the **TripleState** property, a **CheckBox** can also have a[null](vbe-glossary.md) value.
If a  **CheckBox** is[bound](glossary-vba.md) to a[data source](glossary-vba.md), changing the setting changes the value of that source. A disabled  **CheckBox** shows the current value, but is dimmed and does not allow changes to the value from the user interface.
You can also use check boxes inside a group box to select one or more of a group of related items. For example, you can create an order form that contains a list of available items, with a  **CheckBox** preceding each item. The user can select a particular item or items by checking the corresponding **CheckBox**.
The default property of a  **CheckBox** is the **Value** property.
The default event of a  **CheckBox** is the Click event.

 **Note**  The  **ListBox** also lets you put a check mark by selected options. Depending on your application, you can use the **ListBox** instead of using a group of **CheckBox** controls.

 **Related Topics**
[CheckBox Object](http://msdn.microsoft.com/library/03879b09-fe1a-492d-9594-78a82776ecee%28Office.15%29.aspx)

