---
title: LineFormat.EndArrowheadLength Property (Excel)
keywords: vbaxl10.chm110007
f1_keywords:
- vbaxl10.chm110007
ms.prod: excel
api_name:
- Excel.LineFormat.EndArrowheadLength
ms.assetid: e6dd340b-9732-db7e-2efb-7003bca0aea6
ms.date: 06/08/2017
---


# LineFormat.EndArrowheadLength Property (Excel)

Returns or sets the length of the arrowhead at the end of the specified line. Read/write  **[MsoArrowheadLength](http://msdn.microsoft.com/library/e39957f3-ffdd-17fe-dc60-1c3f8c5b14ce%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **EndArrowheadLength**

 _expression_ A variable that represents a **LineFormat** object.


## Remarks





| **MsoArrowheadLength** can be one of these **MsoArrowheadLength** constants.|
| **msoArrowheadLengthMixed**|
| **msoArrowheadShort**|
| **msoArrowheadLengthMedium**|
| **msoArrowheadLong**|

## Example

This example adds a line to  `myDocument`. There's a short, narrow oval on the line's starting point and a long, wide triangle on its end point.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddLine(100, 100, 200, 300).Line 
    .BeginArrowheadLength = msoArrowheadShort 
    .BeginArrowheadStyle = msoArrowheadOval 
    .BeginArrowheadWidth = msoArrowheadNarrow 
    .EndArrowheadLength = msoArrowheadLong 
    .EndArrowheadStyle = msoArrowheadTriangle 
    .EndArrowheadWidth = msoArrowheadWide 
End With
```


## See also


#### Concepts


[LineFormat Object](lineformat-object-excel.md)

