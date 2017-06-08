---
title: CalloutFormat.AutomaticLength Method (Publisher)
keywords: vbapb10.chm2490384
f1_keywords:
- vbapb10.chm2490384
ms.prod: publisher
api_name:
- Publisher.CalloutFormat.AutomaticLength
ms.assetid: 3772ad87-9808-5f25-0b9c-cdd7b1392ca1
ms.date: 06/08/2017
---


# CalloutFormat.AutomaticLength Method (Publisher)

Specifies that the first segment of the callout line (the segment attached to the text callout box) be scaled automatically when the callout is moved.


## Syntax

 _expression_. **AutomaticLength**

 _expression_A variable that represents a  **CalloutFormat** object.


## Remarks

Calling this method sets the  **[AutoLength](calloutformat-autolength-property-publisher.md)** property of the specified object to **msoTrue**.

Use the  **[CustomLength](calloutformat-customlength-method-publisher.md)** method to specify that the first segment of the callout line retain the fixed length returned by the **[Length](calloutformat-length-property-publisher.md)** property whenever the callout is moved. Applies only to callouts whose lines consist of more than one segment (types **msoCalloutThree** and **msoCalloutFour**).


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


