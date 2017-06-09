---
title: CalloutFormat.CustomLength Method (PowerPoint)
keywords: vbapp10.chm559004
f1_keywords:
- vbapp10.chm559004
ms.prod: powerpoint
api_name:
- PowerPoint.CalloutFormat.CustomLength
ms.assetid: 0ee5196b-d3d4-ba8c-ff69-893a92a4ae4d
ms.date: 06/08/2017
---


# CalloutFormat.CustomLength Method (PowerPoint)

Specifies that the first segment of the callout line (the segment attached to the text callout box) retain a fixed length whenever the callout is moved. 


## Syntax

 _expression_. **CustomLength**( **_Length_** )

 _expression_ A variable that represents a **CalloutFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Length_|Required|**Single**|The length of the first segment of the callout, in points.|

### Return Value

Nothing


## Remarks

Use the  **[AutomaticLength](calloutformat-automaticlength-method-powerpoint.md)** method to specify that the first segment of the callout line be scaled automatically whenever the callout is moved. Applies only to callouts whose lines consist of more than one segment (types **msoCalloutThree** and **msoCalloutFour** ).

Applying this method sets the [AutoLength](calloutformat-autolength-property-powerpoint.md)property to  **False** and sets the **Length** property to the value specified for the Length argument.


## Example

This example switches between an automatically scaling first segment and one with a fixed length for the callout line for shape one on  `myDocument`. For the example to work, shape one must be a callout.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).Callout

    If .AutoLength Then

        .CustomLength 50

    Else

        .AutomaticLength

    End If

End With
```


## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-powerpoint.md)

