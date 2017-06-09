---
title: CalloutFormat.AutoAttach Property (PowerPoint)
keywords: vbapp10.chm559008
f1_keywords:
- vbapp10.chm559008
ms.prod: powerpoint
api_name:
- PowerPoint.CalloutFormat.AutoAttach
ms.assetid: bb11ebc3-c84b-9bc0-0bb2-ae30690c7828
ms.date: 06/08/2017
---


# CalloutFormat.AutoAttach Property (PowerPoint)

Determines whether the place where the callout line attaches to the callout text box changes, depending on whether the origin of the callout line (where the callout points to) is to the left or right of the callout text box. Read/write.


## Syntax

 _expression_. **AutoAttach**

 _expression_ A variable that represents an **CalloutFormat** object.


### Return Value

MsoTriState


## Remarks

When the value of this property is  **msoTrue**, the drop value (the vertical distance from the edge of the callout text box to the place where the callout line attaches) is measured from the top of the text box when the text box is to the right of the origin, and it is measured from the bottom of the text box when the text box is to the left of the origin. When the value of this property is **msoFalse**, the drop value is always measured from the top of the text box, regardless of the relative positions of the text box and the origin. Use the **[CustomDrop](calloutformat-customdrop-method-powerpoint.md)** method to set the drop value, and use the **[Drop](calloutformat-drop-property-powerpoint.md)** property to return the drop value.

Setting this property affects a callout only if it has an explicitly set drop value ? that is, if the value of the  **[DropType](calloutformat-droptype-property-powerpoint.md)** property is **msoCalloutDropCustom**. By default, callouts have explicitly set drop values when they're created.

The value of the  **AutoAttach** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The place where the callout line attaches to the callout text box does not change depending on whether the origin of the callout line (where the callout points to) is to the left or right of the callout text box.|
|**msoTrue**| The place where the callout line attaches to the callout text box changes depending on whether the origin of the callout line (where the callout points to) is to the left or right of the callout text box.|

## Example

This example adds two callouts to the first slide. One of the callouts is automatically attached and the other is not. If you change the callout line origin for the automatically attached callout to the right of the attached text box, the position of the text box changes. The callout that is not automatically attached does not display this behavior.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes

    With .AddCallout(msoCalloutTwo, 420, 170, 200, 50)

        .TextFrame.TextRange.Text = "auto-attached"

        .Callout.AutoAttach = msoTrue

    End With

    With .AddCallout(msoCalloutTwo, 420, 350, 200, 50)

        .TextFrame.TextRange.Text = "not auto-attached"

        .Callout.AutoAttach = msoFalse

    End With

End With
```


## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-powerpoint.md)

