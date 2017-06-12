---
title: CalloutFormat.Length Property (Publisher)
keywords: vbapb10.chm2490632
f1_keywords:
- vbapb10.chm2490632
ms.prod: publisher
api_name:
- Publisher.CalloutFormat.Length
ms.assetid: 878fdb7b-fca6-49b6-1ec0-143243ce014c
ms.date: 06/08/2017
---


# CalloutFormat.Length Property (Publisher)

Returns a  **Variant** indicating the length (in points) of the first segment of the callout line (the segment attached to the text callout box) if the **[AutoLength](calloutformat-autolength-property-publisher.md)** property of the specified callout is set to **False**. Otherwise, an error occurs. Read-only.


## Syntax

 _expression_. **Length**

 _expression_A variable that represents a  **CalloutFormat** object.


## Remarks

This property applies only to callouts whose lines consist of more than one segment (types  **msoCalloutThree** and **msoCalloutFour**).

Use the  **[CustomLength](calloutformat-customlength-method-publisher.md)** method to set the value of this property.


## Example

If the first line segment in the callout named co1 has a fixed length, this example specifies that the length of the first line segment in the callout named co2 will also be fixed at that length. For the example to work, both callouts must have multiple-segment lines.


```vb
Dim len1 As Single 
 
With ActiveDocument.Pages(1).Shapes 
 With .Item("co1").Callout 
 If Not .AutoLength Then len1 = .Length 
 End With 
 If len1 Then .Item("co2").Callout _ 
 .CustomLength Length:=len1 
End With
```


