---
title: CalloutFormat.CustomDrop Method (Excel)
keywords: vbaxl10.chm104003
f1_keywords:
- vbaxl10.chm104003
ms.prod: excel
api_name:
- Excel.CalloutFormat.CustomDrop
ms.assetid: d38513f6-1c42-e4b3-7a0f-b8543d59d0ff
ms.date: 06/08/2017
---


# CalloutFormat.CustomDrop Method (Excel)

Sets the vertical distance (in points) from the edge of the text bounding box to the place where the callout line attaches to the text box. This distance is measured from the top of the text box unless the  **[AutoAttach](calloutformat-autoattach-property-excel.md)** property is set to **True** and the text box is to the left of the origin of the callout line (the place that the callout points to), in which case the drop distance is measured from the bottom of the text box.


## Syntax

 _expression_ . **CustomDrop**( **_Drop_** )

 _expression_ A variable that represents a **CalloutFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Drop_|Required| **Single**|The drop distance, in points.|

## Example

This example sets the custom drop distance to 14 points, and specifies that the drop distance always be measured from the top. For the example to work, shape three must be a callout.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3).Callout 
 .CustomDrop 14 
 .AutoAttach = False 
End With
```


## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-excel.md)

