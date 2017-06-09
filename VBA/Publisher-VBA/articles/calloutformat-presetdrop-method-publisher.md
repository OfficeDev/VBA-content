---
title: CalloutFormat.PresetDrop Method (Publisher)
keywords: vbapb10.chm2490387
f1_keywords:
- vbapb10.chm2490387
ms.prod: publisher
api_name:
- Publisher.CalloutFormat.PresetDrop
ms.assetid: a709e54a-d08a-f83c-a0bf-bcdcfe6434cd
ms.date: 06/08/2017
---


# CalloutFormat.PresetDrop Method (Publisher)

Specifies whether the callout line attaches to the top, bottom, or center of the callout text box or whether it attaches at a point that is a specified distance from the top or bottom of the text box.


## Syntax

 _expression_. **PresetDrop**( **_DropType_**)

 _expression_A variable that represents a  **CalloutFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|DropType|Required| **MsoCalloutDropType**|The starting position of the callout line relative to the text bounding box.|

## Remarks

The DropType parameter can be one of the  **MsoCalloutDropType** constants declared in the Microsoft Office type library and shown in the following table.



| **msoCalloutDropBottom**|
| **msoCalloutDropCenter**|
| **msoCalloutDropCustom**|
| **msoCalloutDropTop**|

## Example

This example specifies that the callout line attach to the top of the text bounding box for the first shape in the active publication. For the example to work, the shape must be a callout.


```vb
ActiveDocument.Pages(1).Shapes(1).Callout.PresetDrop DropType:=msoCalloutDropTop
```

This example switches between two preset drops for the first shape one in the active publication. For the example to work, the shape must be a callout.




```vb
With ActiveDocument.Pages(1).Shapes(1).Callout 
 Select Case .DropType 
 Case msoCalloutDropTop 
 .PresetDrop DropType:=msoCalloutDropBottom 
 Case msoCalloutDropBottom 
 .PresetDrop DropType:=msoCalloutDropTop 
 End Select 
End With 

```


