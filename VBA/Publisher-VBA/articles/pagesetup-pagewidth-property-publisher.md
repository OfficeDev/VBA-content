---
title: PageSetup.PageWidth Property (Publisher)
keywords: vbapb10.chm6946822
f1_keywords:
- vbapb10.chm6946822
ms.prod: publisher
api_name:
- Publisher.PageSetup.PageWidth
ms.assetid: 547f2881-d9fa-fa5f-1643-ab08084fb423
ms.date: 06/08/2017
---


# PageSetup.PageWidth Property (Publisher)

Returns or sets a  **Variant** that represents the width of the pages in a publication. Read/write.


## Syntax

 _expression_. **PageWidth**

 _expression_A variable that represents a  **PageSetup** object.


### Return Value

Variant


## Remarks

Numeric values are evaluated as points. String values can be in any unit supported by Microsoft Publisher (for example, "2.5 in"). The valid range of possible values is from zero to the difference between the sheet width and the page width.


## Example

The following example sets a width of eight inches for the pages in the active publication.


```vb
Public Sub PageWidth_Example() 
 ActiveDocument.PageSetup.PageWidth = InchesToPoints(8) 
End Sub
```


