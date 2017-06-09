---
title: CalloutFormat.AutomaticLength Method (Excel)
keywords: vbaxl10.chm104002
f1_keywords:
- vbaxl10.chm104002
ms.prod: excel
api_name:
- Excel.CalloutFormat.AutomaticLength
ms.assetid: e82093e0-7b84-c2c8-8365-6fe05298d55b
ms.date: 06/08/2017
---


# CalloutFormat.AutomaticLength Method (Excel)

Specifies that the first segment of the callout line (the segment attached to the text callout box) be scaled automatically when the callout is moved. Use the  **[CustomLength](calloutformat-customlength-method-excel.md)** method to specify that the first segment of the callout line retain the fixed length returned by the **[Length](calloutformat-length-property-excel.md)** property whenever the callout is moved. Applies only to callouts whose lines consist of more than one segment (types **msoCalloutThree** and **msoCalloutFour** ).


## Syntax

 _expression_ . **AutomaticLength**

 _expression_ A variable that represents a **CalloutFormat** object.


## Remarks

Applying this method sets the  **[AutoLength](calloutformat-autolength-property-excel.md)** property to **True** .


## Example

This example toggles between an automatically scaling first segment and one with a fixed length for the callout line for shape one on  `myDocument`. For the example to work, shape one must be a callout.


```vb
Set myDocument = Worksheets(1) 
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


[CalloutFormat Object](calloutformat-object-excel.md)

