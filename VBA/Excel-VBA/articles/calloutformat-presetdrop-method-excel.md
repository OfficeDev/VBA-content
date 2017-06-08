---
title: CalloutFormat.PresetDrop Method (Excel)
keywords: vbaxl10.chm104005
f1_keywords:
- vbaxl10.chm104005
ms.prod: excel
api_name:
- Excel.CalloutFormat.PresetDrop
ms.assetid: 48d67cad-d93b-2b69-35dd-c3de70340a42
ms.date: 06/08/2017
---


# CalloutFormat.PresetDrop Method (Excel)

Specifies whether the callout line attaches to the top, bottom, or center of the callout text box or whether it attaches at a point that?s a specified distance from the top or bottom of the text box.


## Syntax

 _expression_ . **PresetDrop**( **_DropType_** )

 _expression_ A variable that represents a **CalloutFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DropType_|Required| **[MsoCalloutDropType](http://msdn.microsoft.com/library/0923e0a7-beb6-224f-6a87-85111f58ae3b%28Office.15%29.aspx)**|The starting position of the callout line relative to the text bounding box.|

## Example

This example specifies that the callout line attach to the top of the text bounding box for shape one on  `myDocument`. For the example to work, shape one must be a callout.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(1).Callout.PresetDrop msoCalloutDropTop
```

This example toggles between two preset drops for shape one on  `myDocument`. For the example to work, shape one must be a callout.




```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).Callout 
    If .DropType = msoCalloutDropTop Then 
        .PresetDrop msoCalloutDropBottom 
    ElseIf .DropType = msoCalloutDropBottom Then 
        .PresetDrop msoCalloutDropTop 
    End If 
End With
```


## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-excel.md)

