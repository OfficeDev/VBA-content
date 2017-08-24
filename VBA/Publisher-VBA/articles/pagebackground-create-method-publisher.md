---
title: PageBackground.Create Method (Publisher)
keywords: vbapb10.chm8126469
f1_keywords:
- vbapb10.chm8126469
ms.prod: publisher
api_name:
- Publisher.PageBackground.Create
ms.assetid: a9b699c4-067a-2c68-5f9b-ee7ba0c22cbd
ms.date: 06/08/2017
---


# PageBackground.Create Method (Publisher)

Creates a new  **PageBackground** object for the specified **Page** object.


## Syntax

 _expression_. **Create**

 _expression_A variable that represents a  **PageBackground** object.


## Remarks

Use PageBackground.Exists to test if a page already has a background before trying to create a new one. Returns a "Permission denied' error if a background already exists. 


## Example

The following example tests for the existence of a background on the first page of the active document. If a background does not exist then one is created. 


```vb
If ActiveDocument.Pages(1).Background.Exists = False Then 
 ActiveDocument.Pages(1).Background.Create 
End If
```


