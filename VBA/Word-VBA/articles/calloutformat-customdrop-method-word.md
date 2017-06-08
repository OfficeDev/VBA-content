---
title: CalloutFormat.CustomDrop Method (Word)
keywords: vbawd10.chm163905547
f1_keywords:
- vbawd10.chm163905547
ms.prod: word
api_name:
- Word.CalloutFormat.CustomDrop
ms.assetid: ed727a85-78e4-44f9-a436-f65592cd4be3
ms.date: 06/08/2017
---


# CalloutFormat.CustomDrop Method (Word)

Sets the vertical distance (in points) from the edge of the text bounding box to the place where the callout line attaches to the text box.


## Syntax

 _expression_ . **CustomDrop**( **_Drop_** )

 _expression_ Required. A variable that represents a **[CalloutFormat](calloutformat-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Drop_|Required| **Single**|The drop distance, in points.|

## Remarks

This distance is measured from the top of the text box unless the  **AutoAttach** property is set to **True** and the text box is to the left of the origin of the callout line (the place that the callout points to), in which case the drop distance is measured from the bottom of the text box.

If the  **PresetDrop** method was previously used to set the drop for the specified callout, use the following statement before using the **CustomDrop** method so that the custom drop setting takes effect.




```
PresetDrop msoCalloutDropCustom
```


## Example

This example cancels any preset drop that's been set for the first shape in the active document, sets the custom drop distance to 14 points, and specifies that the drop distance always be measured from the top. For the example to work, the first shape must be a callout.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 

```


```vb
With docActive.Shapes(1).Callout 
 .PresetDrop msoCalloutDropCustom 
 .CustomDrop 14 
 .AutoAttach = False 
End With
```


## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-word.md)

