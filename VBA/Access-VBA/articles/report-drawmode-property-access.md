---
title: Report.DrawMode Property (Access)
keywords: vbaac10.chm13753
f1_keywords:
- vbaac10.chm13753
ms.prod: access
api_name:
- Access.Report.DrawMode
ms.assetid: 773a3c7f-fb59-9614-3363-b417607fbe28
ms.date: 06/08/2017
---


# Report.DrawMode Property (Access)

You can use the  **DrawMode** property to specify how the pen (the color used in drawing) interacts with existing background colors on a report when the **[Line](report-line-method-access.md)**, **[Circle](report-circle-method-access.md)**, or **[Pset](report-pset-method-access.md)** method is used to draw on a report when printing. Read/write **Integer**.


## Syntax

 _expression_. **DrawMode**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **DrawMode** property uses the following settings.



|**Setting**|**Description**|
|:-----|:-----|
|1|Black pen color.|
|2|The inverse of setting 15 (NotMergePen).|
|3|The combination of the colors common to the background color and the inverse of the pen (MaskNotPen).|
|4|The inverse of setting 13 (NotCopyPen).|
|5|The combination of the colors common to both the pen and the inverse of the display (MaskPenNot).|
|6|The inverse of the display color (Invert).|
|7|The combination of the colors in the pen and in the display color, but not in both (XorPen).|
|8|The inverse of setting 9 (NotMaskPen).|
|9|The combination of the colors common to both the pen and the display (MaskPen).|
|10|The inverse of setting 7 (NotXorPen).|
|11|No operation ? the output remains unchanged. In effect, this setting turns drawing off (Nop).|
|12|The combination of the display color and the inverse of the pen color (MergeNotPen).|
|13|(Default) The color specified by the  **ForeColor** property (CopyPen).|
|14|The combination of the pen color and the inverse of the display color (MergePenNot).|
|15|The combination of the pen color and the display color (MergePen).|
|16|White pen color.|
You can set the  **DrawMode** property in an event procedure specified by a section's **OnPrint** property setting.

Use this property to produce visual effects when drawing on a report. Microsoft Access compares each pixel in the draw pattern to the corresponding pixel in the existing background to determine how the drawing appears on the report. For example, setting 7 uses the  **Xor** operator to combine a draw-pattern pixel with a background pixel.


## See also


#### Concepts


[Report Object](report-object-access.md)

