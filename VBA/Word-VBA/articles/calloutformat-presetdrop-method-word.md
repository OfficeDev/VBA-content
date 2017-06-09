---
title: CalloutFormat.PresetDrop Method (Word)
keywords: vbawd10.chm163905549
f1_keywords:
- vbawd10.chm163905549
ms.prod: word
api_name:
- Word.CalloutFormat.PresetDrop
ms.assetid: 3bd6f39f-a5b6-95be-b8de-c60137694d42
ms.date: 06/08/2017
---


# CalloutFormat.PresetDrop Method (Word)

Specifies whether the callout line attaches to the top, bottom, or center of the callout text box or whether it attaches at a point that's a specified distance from the top or bottom of the text box.


## Syntax

 _expression_ . **PresetDrop**( **_DropType_** )

 _expression_ Required. A variable that represents a **[CalloutFormat](calloutformat-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DropType_|Required| **MsoCalloutDropType**|The starting position of the callout line relative to the text bounding box. If you specify  **msoCalloutDropCustom** , the values of the **Drop** and **AutoAttach** properties and the relative positions of the callout text box and callout line origin (the place that the callout points to) are used to determine where the callout line attaches to the text box.|

## Example

This example specifies that the callout line attach to the top of the text bounding box for the first shape on the active document. For the example to work, the first shape must be a callout.


```vb
ActiveDocument.Shapes(1).Callout.PresetDrop msoCalloutDropTop
```

This example switches between two preset drops for the first shape on the active document. For the example to work, the first shape must be a callout.




```vb
With ActiveDocument.Shapes(1).Callout 
 If .DropType = msoCalloutDropTop Then 
 .PresetDrop msoCalloutDropBottom 
 ElseIf .DropType = msoCalloutDropBottom Then 
 .PresetDrop msoCalloutDropTop 
 End If 
End With
```


## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-word.md)

