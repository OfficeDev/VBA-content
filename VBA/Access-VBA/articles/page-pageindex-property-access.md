---
title: Page.PageIndex Property (Access)
keywords: vbaac10.chm12151,vbaac10.chm4455
f1_keywords:
- vbaac10.chm12151,vbaac10.chm4455
ms.prod: access
api_name:
- Access.Page.PageIndex
ms.assetid: 22b71f19-2734-f735-8a64-d02901c598c0
ms.date: 06/08/2017
---


# Page.PageIndex Property (Access)

You can use the  **PageIndex** property to specify or determine the position of a **[Page](page-object-access.md)** object within a **[Pages](pages-object-access.md)** collection. The **PageIndex** property specifies the order in which the pages on a tab control appear. Read/write **Integer**.


## Syntax

 _expression_. **PageIndex**

 _expression_ A variable that represents a **Page** object.


## Remarks

The  **PageIndex** property setting is an **Integer** value between 0 and the **Pages** collection **Count** property setting minus 1.

The  **PageIndex** property can be set in any view.

Changing the value of the  **PageIndex** property changes the location of a **Page** object in the **Pages** collection and visually changes the order of pages on a tab control.


## Example

The following example moves the page named "Notes" on the tab control named "Information" on the "Order Entry" form to the first page.


```vb
Forms("Order Entry").Controls("Information").Pages("Notes").PageIndex = 0 
 

```


## See also


#### Concepts


[Page Object](page-object-access.md)

