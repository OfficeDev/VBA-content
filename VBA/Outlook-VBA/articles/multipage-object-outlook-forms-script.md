---
title: MultiPage Object (Outlook Forms Script)
keywords: olfm10.chm2000570
f1_keywords:
- olfm10.chm2000570
ms.prod: outlook
ms.assetid: ac0fa233-81fe-8a34-4113-6907c6d8f7e2
ms.date: 06/08/2017
---


# MultiPage Object (Outlook Forms Script)

Presents multiple screens of information as a single set.


## Remarks

A  **MultiPage** is useful when you work with a large amount of information that can be sorted into several categories. For example, use a **MultiPage** to display information from an employment application. One page might contain personal information such as name and address; another page might list previous employers; a third page might list references. The **MultiPage** lets you visually combine related information, while keeping the entire record readily accessible.

New pages are added to the right of the currently selected page rather than adjacent to it.

A  **MultiPage** is a control that contains a collection of one or more pages.

Each  **[Page](page-object-outlook-forms-script.md)** of a **MultiPage** is a form that contains its own controls, and as such, can have a unique layout. Typically, the pages in a **MultiPage** have tabs so the user can select the individual pages.

By default, a  **MultiPage** includes two pages, called Page1 and Page2. Each of these is a **Page** object, and together they represent the **[Pages](pages-object-outlook-forms-script.md)** collection of the **MultiPage**. If you add more pages, they become part of the same  **Pages** collection.

The default property for a  **MultiPage** is the **[Value](multipage-value-property-outlook-forms-script.md)** property, which returns the index of the currently active **Page** in the **Pages** collection of the **MultiPage**.

The  **MultiPage** control does not support the **[Click](multipage-click-event-outlook-forms-script.md)** event.


