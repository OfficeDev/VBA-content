---
title: CalloutFormat.CustomDrop Method (PowerPoint)
keywords: vbapp10.chm559003
f1_keywords:
- vbapp10.chm559003
ms.prod: powerpoint
api_name:
- PowerPoint.CalloutFormat.CustomDrop
ms.assetid: 0172ed46-cb73-755a-00c1-cf9c4d29e835
ms.date: 06/08/2017
---


# CalloutFormat.CustomDrop Method (PowerPoint)

Sets the vertical distance (in points) from the edge of the text bounding box to the place where the callout line attaches to the text box. This distance is measured from the top of the text box unless the  **AutoAttach** property is set to **True** and the text box is to the left of the origin of the callout line (the place that the callout points to). In this case the drop distance is measured from the bottom of the text box.


## Syntax

 _expression_. **CustomDrop**( **_Drop_** )

 _expression_ A variable that represents a **CalloutFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Drop_|Required|**Single**|The drop distance, in points.|

### Return Value

Nothing


## Example

This example sets the custom drop distance to 14 points, and specifies that the drop distance always be measured from the top. For the example to work, shape three must be a callout.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3).Callout

    .CustomDrop 14

    .AutoAttach = False

End With
```


## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-powerpoint.md)

