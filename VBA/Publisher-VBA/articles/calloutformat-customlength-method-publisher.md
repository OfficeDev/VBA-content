---
title: CalloutFormat.CustomLength Method (Publisher)
keywords: vbapb10.chm2490386
f1_keywords:
- vbapb10.chm2490386
ms.prod: publisher
api_name:
- Publisher.CalloutFormat.CustomLength
ms.assetid: 855df4af-a02f-fff3-9b12-af886a9788bc
ms.date: 06/08/2017
---


# CalloutFormat.CustomLength Method (Publisher)

Specifies that the first segment of the callout line (the segment attached to the text callout box) retain a fixed length whenever the callout is moved.


## Syntax

 _expression_. **CustomLength**( **_Length_**)

 _expression_A variable that represents a  **CalloutFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Length|Required| **Variant**|The length of the first segment of the callout. Numeric values are evaluated in points; strings can be in any unit supported by Microsoft Publisher (for example, "2.5 in").|

## Remarks

Applying this method sets the  **[AutoLength](calloutformat-autolength-property-publisher.md)** property to  **False** and sets the **[Length](calloutformat-length-property-publisher.md)** property to the value specified for the  **_Length_** argument.

Use the  **[AutomaticLength](calloutformat-automaticlength-method-publisher.md)** method to specify that the first segment of the callout line be scaled automatically whenever the callout is moved. Applies only to callouts whose lines consist of more than one segment (types  **msoCalloutThree** and **msoCalloutFour**).


## Example

This example switches between an automatically-scaling first segment and one with a fixed length for the callout line for the first shape in the active publication. For the example to work, this shape must be a callout.


```vb
With ActiveDocument.Pages(1).Shapes(1).Callout 
 If .AutoLength Then 
 .CustomLength Length:=50 
 Else 
 .AutomaticLength 
 End If 
End With
```


