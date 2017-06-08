---
title: CalloutFormat.AutoAttach Property (Excel)
keywords: vbaxl10.chm104008
f1_keywords:
- vbaxl10.chm104008
ms.prod: excel
api_name:
- Excel.CalloutFormat.AutoAttach
ms.assetid: 80f5bf63-072d-1245-d564-1b54af0f85b5
ms.date: 06/08/2017
---


# CalloutFormat.AutoAttach Property (Excel)

 **True** if the place where the callout line attaches to the callout text box changes depending on whether the origin of the callout line (where the callout points to) is to the left or right of the callout text box. Read/write **[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **AutoAttach**

 _expression_ A variable that represents a **CalloutFormat** object.


## Remarks



| **MsoTriState** can be one of these **MsoTriState** constants.|
| **msoCTrue**|
| **msoFalse**|
| **msoTriStateMixed**|
| **msoTriStateToggle**|
| **msoTrue** . The place where the callout line attaches to the callout text box changes depending on whether the origin of the callout line (where the callout points to) is to the left or right of the callout text box.|
When the value of this property is  **True** , the drop value (the vertical distance from the edge of the callout text box to the place where the callout line attaches) is measured from the top of the text box when the text box is to the right of the origin, and it's measured from the bottom of the text box when the text box is to the left of the origin. When the value of this property is **False** , the drop value is always measured from the top of the text box, regardless of the relative positions of the text box and the origin. Use the **[CustomDrop](calloutformat-customdrop-method-excel.md)** method to set the drop value, and use the **[Drop](calloutformat-drop-property-excel.md)** property to return the drop value.

Setting this property affects a callout only if it has an explicitly set drop value — that is, if the value of the  **[DropType](calloutformat-droptype-property-excel.md)** property is **msoCalloutDropCustom** . By default, callouts have explicitly set drop values when they're created.


## Example

This example adds two callouts to  `myDocument`. If you drag the text box for each of these callouts to the left of the callout line origin, the place on the text box where the callout line attaches will change for the automatically attached callout.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes 
    With .AddCallout(msoCalloutTwo, 420, 170, 200, 50) 
        .TextFrame.Characters.Text = "auto-attached" 
        .Callout.AutoAttach = True 
    End With 
    With .AddCallout(msoCalloutTwo, 420, 350, 200, 50) 
        .TextFrame.Characters.Text = "not auto-attached" 
        .Callout.AutoAttach = False 
    End With 
End With
```


## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-excel.md)

