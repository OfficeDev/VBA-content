---
title: CalloutFormat.DropType Property (Publisher)
keywords: vbapb10.chm2490630
f1_keywords:
- vbapb10.chm2490630
ms.prod: publisher
api_name:
- Publisher.CalloutFormat.DropType
ms.assetid: fd4ec192-0732-e860-4ff8-e305aa0d90a9
ms.date: 06/08/2017
---


# CalloutFormat.DropType Property (Publisher)

Returns an  **MsoCalloutDropType** constant indicating where the callout line attaches to the callout text box. Read-only.


## Syntax

 _expression_. **DropType**

 _expression_A variable that represents a  **CalloutFormat** object.


### Return Value

MsoCalloutDropType


## Remarks

The  **DropType** property value can be one of the ** [MsoCalloutDropType](http://msdn.microsoft.com/library/0923e0a7-beb6-224f-6a87-85111f58ae3b%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.

If the callout drop type is  **msoCalloutDropCustom**, the values of the  **[Drop](calloutformat-drop-property-publisher.md)** and  **[AutoAttach](calloutformat-autoattach-property-publisher.md)** properties and the relative positions of the callout text box and callout line origin (where the callout points) are used to determine where the callout line attaches to the text box.

Use the  **[PresetDrop](calloutformat-presetdrop-method-publisher.md)** method to set the value of this property.


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


