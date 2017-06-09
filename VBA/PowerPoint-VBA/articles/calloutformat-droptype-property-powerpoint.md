---
title: CalloutFormat.DropType Property (PowerPoint)
keywords: vbapp10.chm559012
f1_keywords:
- vbapp10.chm559012
ms.prod: powerpoint
api_name:
- PowerPoint.CalloutFormat.DropType
ms.assetid: 993a7cb5-afc6-0683-d8f1-5b71633f07bf
ms.date: 06/08/2017
---


# CalloutFormat.DropType Property (PowerPoint)

Returns a value that indicates where the callout line attaches to the callout text box. Read-only.


## Syntax

 _expression_. **DropType**

 _expression_ A variable that represents a **CalloutFormat** object.


### Return Value

MsoCalloutDropType


## Remarks

If the callout drop type is  **msoCalloutDropCustom**, the values of the[Drop](calloutformat-drop-property-powerpoint.md)and  **[AutoAttach](calloutformat-autoattach-property-powerpoint.md)** properties and the relative positions of the callout text box and callout line origin (the place that the callout points to) are used to determine where the callout line attaches to the text box.

This property is read-only. Use the  **[PresetDrop](calloutformat-presetdrop-method-powerpoint.md)** method to set the value of this property.

The value returned by the  **DropType** property can be one of these **MsoCalloutDropType** constants.


||
|:-----|
|**msoCalloutDropBottom**|
|**msoCalloutDropCenter**|
|**msoCalloutDropCustom**|
|**msoCalloutDropMixed**|
|**msoCalloutDropTop**|

## Example

This example checks to determine whether shape three on  `myDocument` is a callout with a custom drop. If it is, the code replaces the custom drop with one of two preset drops, depending on whether the custom drop value is greater than or less than half the height of the callout text box.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3)

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


[CalloutFormat Object](calloutformat-object-powerpoint.md)

