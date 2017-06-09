---
title: Items.GetPrevious Method (Outlook)
keywords: vbaol11.chm67
f1_keywords:
- vbaol11.chm67
ms.prod: outlook
api_name:
- Outlook.Items.GetPrevious
ms.assetid: 5dde47f8-2bd8-fdbe-d6e7-b1381e8a97a6
ms.date: 06/08/2017
---


# Items.GetPrevious Method (Outlook)

Returns the previous object in the collection. 


## Syntax

 _expression_ . **GetPrevious**

 _expression_ A variable that represents an **Items** object.


### Return Value

An  **Object** value that represents the previous object contained by the collection.


## Remarks

It returns  **Nothing** if no previous object exists, for example, if already positioned at the beginning of the collection. To ensure correct operation of the **[GetFirst](items-getfirst-method-outlook.md)** , **[GetLast](items-getlast-method-outlook.md)** , **[GetNext](items-getnext-method-outlook.md)** , and **GetPrevious** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection, and call **GetLast** before calling **GetPrevious** . To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## See also


#### Concepts


[Items Object](items-object-outlook.md)

