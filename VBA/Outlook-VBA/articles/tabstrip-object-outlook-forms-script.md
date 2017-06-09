---
title: TabStrip Object (Outlook Forms Script)
keywords: olfm10.chm2000660
f1_keywords:
- olfm10.chm2000660
ms.prod: outlook
ms.assetid: 643c896a-2304-42f3-f5e9-0feee6d22364
ms.date: 06/08/2017
---


# TabStrip Object (Outlook Forms Script)

Presents a set of related controls as a visual group.


## Remarks

You can use a  **TabStrip** to view different sets of information for related controls.

A  **TabStrip** is a control that contains a collection of one or more tabs.

Each  **[Tab](tab-object-outlook-forms-script.md)** of a **TabStrip** is a separate object that users can select. Visually, a **TabStrip** also includes a client area that all the tabs in the **TabStrip** share.

By default, a  **TabStrip** includes two pages, called Tab1 and Tab2. Each of these is a **Tab** object, and together they represent the **[Tabs](tabs-object-outlook-forms-script.md)** collection of the **TabStrip**. If you add more pages, they become part of the same  **Tabs** collection.

For example, the controls might represent information about a daily schedule for a group of individuals, with each set of information corresponding to a different individual in the group. Set the title of each tab to show one individual's name. Then, you can write code that, after you click a tab, updates the controls to show information about the person identified on the tab.

The  **TabStrip** is implemented as a container of a **Tabs** collection, which in turn contains a group of **Tab** objects. The **TabStrip** control does not support the **Click** event.

The default property for a  **TabStrip** is the **[SelectedItem](tabstrip-selecteditem-property-outlook-forms-script.md)** property.


