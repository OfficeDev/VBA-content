---
title: Shapes.BuildFreeform Method (Word)
keywords: vbawd10.chm161415188
f1_keywords:
- vbawd10.chm161415188
ms.prod: word
api_name:
- Word.Shapes.BuildFreeform
ms.assetid: 760fe720-3fbc-16a1-c5b3-b78502dbf670
ms.date: 06/08/2017
---


# Shapes.BuildFreeform Method (Word)

Builds a freeform object.


## Syntax

 _expression_ . **BuildFreeform**( **_EditingType_** , **_X1_** , **_Y1_** )

 _expression_ Required. A variable that represents a **[Shapes](shapes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EditingType_|Required| **MsoEditingType**|The editing property of the first node.|
| _X1_|Required| **Single**|The position (in points) of the first node in the freeform drawing relative to the left edge of the document.|
| _Y1_|Required| **Single**|The position (in points) of the first node in the freeform drawing relative to the top edge of the document.|

### Return Value

 **[FreeformBuilder](freeformbuilder-object-word.md)**


## Remarks

Use the  **AddNodes** method to add segments to the freeform. After you have added at least one segment to the freeform, you can use the **ConvertToShape** method to convert the **[FreeformBuilder](freeformbuilder-object-word.md)** object into a **Shape** object that has the geometric description you've defined in the **[FreeformBuilder](freeformbuilder-object-word.md)** object.


## Example

This example adds a freeform with five vertices to the active document.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 

```


```vb
With docActive.Shapes.BuildFreeform(msoEditingCorner, 360, 200) 
 .AddNodes msoSegmentCurve, msoEditingCorner, _ 
 380, 230, 400, 250, 450, 300 
 .AddNodes msoSegmentCurve, msoEditingAuto, 480, 200 
 .AddNodes msoSegmentLine, msoEditingAuto, 480, 400 
 .AddNodes msoSegmentLine, msoEditingAuto, 360, 200 
 .ConvertToShape 
End With
```


## See also


#### Concepts


[Shapes Collection Object](shapes-object-word.md)

