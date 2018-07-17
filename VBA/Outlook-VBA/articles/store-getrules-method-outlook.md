---
title: Store.GetRules Method (Outlook)
keywords: vbaol11.chm810
f1_keywords:
- vbaol11.chm810
ms.prod: outlook
api_name:
- Outlook.Store.GetRules
ms.assetid: 06048799-e162-68f9-17c2-d80c25e2c55e
ms.date: 06/08/2017
---


# Store.GetRules Method (Outlook)

Returns a  **[Rules](rules-object-outlook.md)** collection object that contains the **[Rule](rule-object-outlook.md)** objects defined for the current session.


## Syntax

 _expression_ . **GetRules**

 _expression_ A variable that represents a **Store** object.


### Return Value

A  **Rules** collection object that represents the set of **Rules** defined for the current session.


## Remarks

Calling  **GetRules** can be an expensive operation in terms of performance on slow connections to an Exchange server.

The order of the  **Rule** objects in the collection returned from **GetRules** follows that of **[Rule.ExecutionOrder](rule-executionorder-property-outlook.md)** with **ExecutionOrder** equal 1 being the first **Rule** in the collection and **ExecutionOrder** equal **[Rules.Count](rules-count-property-outlook.md)** being the last **Rule** in the collection.


## See also


#### Concepts


[Store Object](store-object-outlook.md)

