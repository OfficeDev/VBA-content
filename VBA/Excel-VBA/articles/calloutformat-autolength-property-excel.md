---
title: CalloutFormat.AutoLength Property (Excel)
keywords: vbaxl10.chm104009
f1_keywords:
- vbaxl10.chm104009
ms.prod: excel
api_name:
- Excel.CalloutFormat.AutoLength
ms.assetid: aadce7bf-e4b3-b56d-8a10-cf8183282149
ms.date: 06/08/2017
---


# CalloutFormat.AutoLength Property (Excel)

Applies only to callouts whose lines consist of more than one segment (types  **msoCalloutThree** and **msoCalloutFour** ). Read/write **[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **AutoLength**

 _expression_ A variable that represents a **CalloutFormat** object.


## Remarks



| **MsoTriState** can be one of these **MsoTriState** constants.|
| **msoCTrue**|
| **msoFalse** . The first segment of the callout retains the fixed length specified by the **Length** property whenever the callout is moved.|
| **msoTriStateMixed**|
| **msoTriStateToggle**|
| **msoTrue** . The first segment of the callout line (the segment attached to the text callout box) is scaled automatically whenever the callout is moved.|
This property is read-only. Use the  **[AutomaticLength](calloutformat-automaticlength-method-excel.md)** method to set this property to **msoTrue** , and use the **[CustomLength](calloutformat-customlength-method-excel.md)** method to set this property to **mosFalse** .


## Example

This example toggles between an automatically scaling first segment and one with a fixed length for the callout line for shape one on  `myDocument`. For the example to work, shape one must be a callout.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).Callout 
    If .AutoLength Then 
        .CustomLength 50 
    Else 
        .AutomaticLength 
    End If 
End With
```


## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-excel.md)

