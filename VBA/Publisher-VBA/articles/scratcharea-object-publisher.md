---
title: ScratchArea Object (Publisher)
keywords: vbapb10.chm1245183
f1_keywords:
- vbapb10.chm1245183
ms.prod: publisher
api_name:
- Publisher.ScratchArea
ms.assetid: 41856866-c1d8-2550-1b4c-28886ed2b714
ms.date: 06/08/2017
---


# ScratchArea Object (Publisher)

Represents the area outside the boundaries of publication pages where layout elements may be stored with no effect on publication output.
 


## Example

Use the  **[ScratchArea](document-scratcharea-property-publisher.md)** property of the **Document** object to return a scratch area. Use the **Shapes** property of the **ScratchArea** object to return the collection of shapes that are currently on a scratch area.
 

 

 

 
This example assigns the first shape on the scratch area of the active document to a variable.
 

 



```
Dim saPage As ScratchArea 
Dim objFirst As Object 
 
saPage = Application.ActiveDocument.ScratchArea 
objFirst = saPage.Shapes(1)
```


## Properties



|**Name**|
|:-----|
|[Application](scratcharea-application-property-publisher.md)|
|[Parent](scratcharea-parent-property-publisher.md)|
|[Shapes](scratcharea-shapes-property-publisher.md)|

