---
title: CalloutFormat.CustomLength Method (Excel)
keywords: vbaxl10.chm104004
f1_keywords:
- vbaxl10.chm104004
ms.prod: excel
api_name:
- Excel.CalloutFormat.CustomLength
ms.assetid: 8c5034f9-32ca-6e34-be59-51e0cd8c8374
ms.date: 06/08/2017
---


# CalloutFormat.CustomLength Method (Excel)

Specifies that the first segment of the callout line (the segment attached to the text callout box) retain a fixed length whenever the callout is moved. Use the  **[AutomaticLength](calloutformat-automaticlength-method-excel.md)** method to specify that the first segment of the callout line be scaled automatically whenever the callout is moved. Applies only to callouts whose lines consist of more than one segment (types **msoCalloutThree** and **msoCalloutFour** ).


## Syntax

 _expression_ . **CustomLength**( **_Length_** )

 _expression_ A variable that represents a **CalloutFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Length_|Required| **Single**|The length of the first segment of the callout, in points.|

## Remarks

Applying this method sets the  **[AutoLength](calloutformat-autolength-property-excel.md)** property to **False** and sets the **[Length](calloutformat-length-property-excel.md)** property to the value specified for the _Length_ argument.


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

