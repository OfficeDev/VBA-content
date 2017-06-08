---
title: CalloutFormat.Drop Property (Excel)
keywords: vbaxl10.chm104011
f1_keywords:
- vbaxl10.chm104011
ms.prod: excel
api_name:
- Excel.CalloutFormat.Drop
ms.assetid: fd1845fb-bdef-aa9e-5e49-a6c2fd6e2cb6
ms.date: 06/08/2017
---


# CalloutFormat.Drop Property (Excel)

For callouts with an explicitly set drop value, this property returns the vertical distance (in points) from the edge of the text bounding box to the place where the callout line attaches to the text box. Read-only  **Single** .


## Syntax

 _expression_ . **Drop**

 _expression_ A variable that represents a **CalloutFormat** object.


## Remarks

This distance is measured from the top of the text box unless the  **AutoAttach** property is set to **True** and the text box is to the left of the origin of the callout line (the place that the callout points to), in which case the drop distance is measured from the bottom of the text box.

Use the  **[CustomDrop](calloutformat-customdrop-method-excel.md)** method to set the value of this property.

The value of this property accurately reflects the position of the callout line attachment to the text box only if the callout has an explicitly set drop value — that is, if the value of the  **[DropType](calloutformat-droptype-property-excel.md)** property is **msoCalloutDropCustom** .


## Example

This example replaces the custom drop for shape one on  `myDocument` with one of two preset drops, depending on whether the custom drop value is greater than or less than half the height of the callout text box. For the example to work, shape one must be a callout.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).Callout 
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


[CalloutFormat Object](calloutformat-object-excel.md)

