---
title: Shape.Hyperlink Property (Publisher)
keywords: vbapb10.chm2228323
f1_keywords:
- vbapb10.chm2228323
ms.prod: publisher
api_name:
- Publisher.Shape.Hyperlink
ms.assetid: 0990ab32-b4a3-6c89-cb9f-8f8c64ef804f
ms.date: 06/08/2017
---


# Shape.Hyperlink Property (Publisher)

Returns a  **[Hyperlink](hyperlink-object-publisher.md)** object representing the hyperlink associated with the specified shape.


## Syntax

 _expression_. **Hyperlink**

 _expression_A variable that represents a  **Shape** object.


## Example

This example sets shape one on page one in the active publication to jump to the specified Web site when the shape is clicked.


```vb
Dim hypTemp As Hyperlink 
 
Set hypTemp = ActiveDocument.Pages(1).Shapes(1).Hyperlink 
 
hypTemp.Address = "http://www.tailspintoys.com/"
```


