---
title: CalloutFormat.Drop Property (Publisher)
keywords: vbapb10.chm2490629
f1_keywords:
- vbapb10.chm2490629
ms.prod: publisher
api_name:
- Publisher.CalloutFormat.Drop
ms.assetid: 7878a6a6-9c7c-dfd0-ef1b-d56a5aab6a18
ms.date: 06/08/2017
---


# CalloutFormat.Drop Property (Publisher)

For callouts with an explicitly set drop value, this property returns the vertical distance from the edge of the text bounding box to the place where the callout line attaches to the text box. This distance is measured from the top of the text box unless the  **AutoAttach** property is set to **True** and the text box is to the left of the origin of the callout line (where the callout points). In this case, the drop distance is measured from the bottom of the text box. Read-only **Variant**.


## Syntax

 _expression_. **Drop**

 _expression_A variable that represents a  **CalloutFormat** object.


### Return Value

Variant


## Remarks

Numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").

Use the  **[CustomDrop](calloutformat-customdrop-method-publisher.md)** method to set the value of this property.

The value of this property accurately reflects the position of the callout line attachment to the text box only if the callout has an explicitly set drop value — that is, if the value of the  **[DropType](calloutformat-droptype-property-publisher.md)** property is  **msoCalloutDropCustom**.


## Example

This example replaces the custom drop for the first shape in the active publication with one of two preset drops, depending on whether the custom drop value is greater than or less than half the height of the callout text box. For the example to work, the shape must be a callout.


```vb
With ActiveDocument.Pages(1).Shapes(1).Callout 
 If .DropType = msoCalloutDropCustom Then 
 If .Drop < .Parent.Height / 2 Then 
 .PresetDrop DropType:=msoCalloutDropTop 
 Else 
 .PresetDrop DropType:=msoCalloutDropBottom 
 End If 
 End If 
End With 

```


