---
title: Pages.Count Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 67070350-7267-979c-8205-c64bc3e147b4
ms.date: 06/08/2017
---


# Pages.Count Property (Outlook Forms Script)

Returns a  **Long** that represents the number of objects in a collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_A variable that represents a  **Pages** object.


## Remarks

The  **Count** property is read only.

Note that the index value for the first page of a collection is zero, the value for the second page is one, and so on. For example, if a  **[MultiPage](multipage-object-outlook-forms-script.md)** contains two pages, the indexes of the pages are 0 and 1, and the value of **Count** is 2.


## See also


#### Concepts


 [Pages Object](pages-object-outlook-forms-script.md)

