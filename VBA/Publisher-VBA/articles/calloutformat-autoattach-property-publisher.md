---
title: CalloutFormat.AutoAttach Property (Publisher)
keywords: vbapb10.chm2490626
f1_keywords:
- vbapb10.chm2490626
ms.prod: publisher
api_name:
- Publisher.CalloutFormat.AutoAttach
ms.assetid: 893303d8-97fe-9eea-8d6e-d9110c75ee84
ms.date: 06/08/2017
---


# CalloutFormat.AutoAttach Property (Publisher)

Returns or sets an  **MsoTriState**constant indicating whether the place where the callout line attaches to the callout text box changes depending on whether the origin of the callout line (where the callout points) is to the left or right of the callout text box. Read/write.


## Syntax

 _expression_. **AutoAttach**

 _expression_A variable that represents a  **CalloutFormat** object.


### Return Value

MsoTriState


## Remarks

The  **AutoAttach** property value can be one of the ** [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.

When the value of this property is  **msoTrue**, the drop value (the vertical distance from the edge of the callout text box to the place where the callout line attaches) is measured from the top of the text box when the text box is to the right of the origin, and it is measured from the bottom of the text box when the text box is to the left of the origin. When the value of this property is  **msoFalse**, the drop value is always measured from the top of the text box, regardless of the relative positions of the text box and the origin. Use the  [CustomDrop](calloutformat-customdrop-method-publisher.md)method to set the drop value, and use the  [Drop](calloutformat-drop-property-publisher.md)property to return the drop value.

Setting this property affects a callout only if it has an explicitly set drop value—that is, if the value of the  [DropType](calloutformat-droptype-property-publisher.md)property is  **msoCalloutDropCustom**. By default, callouts have explicitly set drop values when they are created.


## Example

This example adds two callouts to the first page. One of the callouts is automatically attached and the other is not. If you change the callout line origin for the automatically attached callout to the right of the attached text box, the position of the text box changes. The callout that is not automatically attached does not display this behavior.


```vb
With ActivePublication.Pages(1).Shapes 
 With .AddCallout(Type:=msoCalloutTwo, _ 
 Left:=420, Top:=170, Width:=200, Height:=50) 
 .TextFrame.TextRange.Text = "auto-attached" 
 .Callout.AutoAttach = msoTrue 
 End With 
 With .AddCallout(Type:=msoCalloutTwo, _ 
 Left:=420, Top:=350, Width:=200, Height:=50) 
 .TextFrame.TextRange.Text = "not auto-attached" 
 .Callout.AutoAttach = msoFalse 
 End With 
End With 

```


