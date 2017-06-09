---
title: Tab.Index Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 2cacd35e-edd4-6733-e932-a05114134754
ms.date: 06/08/2017
---


# Tab.Index Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies the position of a **[Tab](tab-object-outlook-forms-script.md)** object within a **[Tabs](tabs-object-outlook-forms-script.md)** collection. Read/write.


## Syntax

 _expression_. **Index**

 _expression_A variable that represents a  **Tab** object.


## Remarks

The  **Index** property specifies the order in which tabs appear. Changing the value of **Index** visually changes the order of tabs on a **[TabStrip](tabstrip-object-outlook-forms-script.md)**. The index value for the first tab is zero, the index value of the second tab is one, and so on.

In a  **MultiPage**,  **Index** refers to a **Page** as well as the page's **Tab**. In a  **TabStrip**,  **Index** refers to the tab only.


