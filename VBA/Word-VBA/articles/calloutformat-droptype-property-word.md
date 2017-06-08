---
title: CalloutFormat.DropType Property (Word)
keywords: vbawd10.chm163905642
f1_keywords:
- vbawd10.chm163905642
ms.prod: word
api_name:
- Word.CalloutFormat.DropType
ms.assetid: cf26ef87-7c56-5859-75fc-dfee7bd0efd1
ms.date: 06/08/2017
---


# CalloutFormat.DropType Property (Word)

Returns a value that indicates where the callout line attaches to the callout text box. Read-only  **MsoCalloutDropType** .


## Syntax

 _expression_ . **DropType**

 _expression_ Required. A variable that represents a **[CalloutFormat](calloutformat-object-word.md)** object.


## Remarks

If the callout drop type is  **msoCalloutDropCustom** , the values of the **Drop** and **AutoAttach** properties and the relative positions of the callout text box and callout line origin (the place that the callout points to) are used to determine where the callout line attaches to the text box.

This property is read-only. Use the  **PresetDrop** method to set the value of this property.


## Example

This example checks to determine whether the third shape on the active document is a callout with a custom drop. If it is, the code replaces the custom drop with one of two preset drops, depending on whether the custom drop value is greater than or less than half the height of the callout text box.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 
 
With docActive.Shapes(3) 
 If .Type = msoCallout Then 
 With .Callout 
 If .DropType = msoCalloutDropCustom Then 
 If .Drop < .Parent.Height / 2 Then 
 .PresetDrop msoCalloutDropTop 
 Else 
 .PresetDrop msoCalloutDropBottom 
 End If 
 End If 
 End With 
 End If 
End With
```


## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-word.md)

