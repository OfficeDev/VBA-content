---
title: Shapes.AddCallout Method (Excel)
keywords: vbaxl10.chm638077
f1_keywords:
- vbaxl10.chm638077
ms.prod: excel
api_name:
- Excel.Shapes.AddCallout
ms.assetid: b98ea95d-210b-34cc-c999-e7ce0a3e3a72
ms.date: 06/08/2017
---


# Shapes.AddCallout Method (Excel)

 Creates a borderless line callout. Returns a **[Shape](shape-object-excel.md)** object that represents the new callout.


## Syntax

 _expression_ . **AddCallout**( **_Type_** , **_Left_** , **_Top_** , **_Width_** , **_Height_** )

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[MsoCalloutType](http://msdn.microsoft.com/library/65548284-0241-f013-ea54-93099fdbf1cc%28Office.15%29.aspx)**|The type of callout line.|
| _Left_|Required| **Single**|The position (in points) of the upper-left corner of the callout's bounding box relative to the upper-left corner of the document.|
| _Top_|Required| **Single**|The position (in points) of the upper-left corner of the callout's bounding box relative to the upper-left corner of the document.|
| _Width_|Required| **Single**|The width of the callout's bounding box, in points.|
| _Height_|Required| **Single**|The height of the callout's bounding box, in points.|

### Return Value

Shape


## Remarks



| **MsoCalloutType** can be one of these **MsoCalloutType** constants.|
| **msoCalloutOne** . A single-segment callout line that can be either horizontal or vertical.|
| **msoCalloutTwo** . A single-segment callout line that rotates freely.|
| **msoCalloutMixed** .|
| **msoCalloutThree** . A two-segment line.|
| **msoCalloutFour** . A three-segment line.|
You can insert a greater variety of callouts by using the  **[AddShape](shapes-addshape-method-excel.md)** method.


## Example

This example adds a borderless callout with a freely rotating one-segment callout line to  `myDocument` and then sets the callout angle to 30 degrees.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.AddCallout(Type:=msoCalloutTwo, _ 
    Left:=50, Top:=50, Width:=200, Height:=100) _ 
    .Callout.Angle = msoCalloutAngle30
```


## See also


#### Concepts


[Shapes Object](shapes-object-excel.md)

