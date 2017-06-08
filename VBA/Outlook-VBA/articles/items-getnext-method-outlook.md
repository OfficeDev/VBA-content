---
title: Items.GetNext Method (Outlook)
keywords: vbaol11.chm66
f1_keywords:
- vbaol11.chm66
ms.prod: outlook
api_name:
- Outlook.Items.GetNext
ms.assetid: 01c49c21-d9f9-37c4-8c64-ff8e2b1f9462
ms.date: 06/08/2017
---


# Items.GetNext Method (Outlook)

Returns the next object in the collection. 


## Syntax

 _expression_ . **GetNext**

 _expression_ A variable that represents an **Items** object.


### Return Value

An  **Object** value that represents the next object contained by the collection.


## Remarks

It returns  **Nothing** if no next object exists, for example, if already positioned at the end of the collection. To ensure correct operation of the **[GetFirst](items-getfirst-method-outlook.md)** , **[GetLast](items-getlast-method-outlook.md)** , **GetNext** , and **[GetPrevious](items-getprevious-method-outlook.md)** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection, and call **GetLast** before calling **GetPrevious** . To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## See also


#### Concepts


[Items Object](items-object-outlook.md)

