---
title: CalloutFormat.Gap Property (Excel)
keywords: vbaxl10.chm104013
f1_keywords:
- vbaxl10.chm104013
ms.prod: excel
api_name:
- Excel.CalloutFormat.Gap
ms.assetid: 6f50eb69-23f8-a9a1-e0cf-16caf76f3263
ms.date: 06/08/2017
---


# CalloutFormat.Gap Property (Excel)

Returns or sets the horizontal distance (in points) between the end of the callout line and the text bounding box. Read/write  **Single** .


## Syntax

 _expression_ . **Gap**

 _expression_ A variable that represents a **CalloutFormat** object.


## Example

This example sets the distance between the callout line and the text bounding box to 3 points for shape one on  `myDocument`. For the example to work, shape one must be a callout.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(1).Callout.Gap = 3
```


## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-excel.md)

