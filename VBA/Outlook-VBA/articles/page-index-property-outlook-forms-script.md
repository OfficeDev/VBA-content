---
title: Page.Index Property (Outlook Forms Script)
keywords: olfm10.chm2001280
f1_keywords:
- olfm10.chm2001280
ms.prod: outlook
ms.assetid: 91e67439-ea23-9ac8-6065-31af7be0b303
ms.date: 06/08/2017
---


# Page.Index Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies the position of a **[Page](page-object-outlook-forms-script.md)** object in a **[Pages](pages-object-outlook-forms-script.md)** collection. Read/write.


## Syntax

 _expression_. **Index**

 _expression_A variable that represents a  **Page** object.


## Remarks

The  **Index** property specifies the order in which tabs appear. Changing the value of **Index** visually changes the order of pages in a **[MultiPage](multipage-object-outlook-forms-script.md)**. The index value for the first page is zero, the index value of the second page is one, and so on.

In a  **MultiPage**,  **Index** refers to a **Page** as well as the page's **[Tab](tab-object-outlook-forms-script.md)**.


