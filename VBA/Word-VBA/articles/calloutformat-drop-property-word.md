---
title: CalloutFormat.Drop Property (Word)
keywords: vbawd10.chm163905641
f1_keywords:
- vbawd10.chm163905641
ms.prod: word
api_name:
- Word.CalloutFormat.Drop
ms.assetid: e68a15a5-a976-bb70-f11f-f7eec126bb0a
ms.date: 06/08/2017
---


# CalloutFormat.Drop Property (Word)

Returns the vertical distance (in points) from the edge of the text bounding box to the place where the callout line attaches to the text box. Read-only  **Single** .


## Syntax

 _expression_ . **Drop**

 _expression_ A variable that represents a **[CalloutFormat](calloutformat-object-word.md)** object.


## Remarks

The  **Drop** property applies to callouts with an explicitly set drop value. This distance is measured from the top of the text box unless the text box is to the left of the origin of the callout line (the place that the callout points to), in which case the drop distance is measured from the bottom of the text box.

Use the  **[CustomDrop](calloutformat-customdrop-method-word.md)** method to set the value of this property.

The value of this property accurately reflects the position of the callout line attachment to the text box only if the callout has an explicitly set drop value — that is, if the value of the  **[DropType](calloutformat-droptype-property-word.md)** property is **msoCalloutDropCustom** . Use the statement `PresetDrop msoCalloutCustomDrop` to set the **DropType** property to **msoCalloutDropCustom** .


## Example

This example replaces the custom drop for the first shape on the active document with one of two preset drops, depending on whether the custom drop value is greater than or less than half the height of the callout text box. For the example to work, the first shape must be a callout.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 
 
With docActive.Shapes(1).Callout 
 If .DropType = msoCalloutDropCustom Then 
 If .Drop < .Parent.Height / 2 Then 
 .PresetDrop msoCalloutDropTop 
 Else 
 .PresetDrop msoCalloutDropBottom 
 End If 
 End If 
End With
```


## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-word.md)

