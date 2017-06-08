---
title: CalloutFormat.DropType Property (Excel)
keywords: vbaxl10.chm104012
f1_keywords:
- vbaxl10.chm104012
ms.prod: excel
api_name:
- Excel.CalloutFormat.DropType
ms.assetid: ab947fa4-4af9-e491-f62d-e0ca036e1892
ms.date: 06/08/2017
---


# CalloutFormat.DropType Property (Excel)

Returns a value that indicates where the callout line attaches to the callout text box. Read-only  **[MsoCalloutDropType](http://msdn.microsoft.com/library/0923e0a7-beb6-224f-6a87-85111f58ae3b%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **DropType**

 _expression_ A variable that represents a **CalloutFormat** object.


## Remarks



| **MsoCalloutDropType** can be one of these **MsoCalloutDropType** constants.|
| **msoCalloutDropCenter**|
| **msoCalloutDropMixed**|
| **msoCalloutDropBottom**|
| **msoCalloutDropCustom**|
| **msoCalloutDropTop**|
If the callout drop type is  **msoCalloutDropCustom** , the values of the **[Drop](calloutformat-drop-property-excel.md)** and **[AutoAttach](calloutformat-autoattach-property-excel.md)** properties and the relative positions of the callout text box and callout line origin (the place that the callout points to) are used to determine where the callout line attaches to the text box.

This property is read-only. Use the  **[PresetDrop](calloutformat-presetdrop-method-excel.md)** method to set the value of this property.


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

