---
title: Report.ScaleMode Property (Access)
keywords: vbaac10.chm13745
f1_keywords:
- vbaac10.chm13745
ms.prod: access
api_name:
- Access.Report.ScaleMode
ms.assetid: e3955e48-80bb-989e-2992-cd5a541b468b
ms.date: 06/08/2017
---


# Report.ScaleMode Property (Access)

You can use the  **ScaleMode** property in Visual Basic to specify the unit of measurement for coordinates on a page when the **[Circle](report-circle-method-access.md)**, **[Line](report-line-method-access.md)**, **[Pset](report-pset-method-access.md)**, or **[Print](report-print-method-access.md)** method is used while a report is previewed or printed, or its output is saved to a file. Read/write **Integer**.


## Syntax

 _expression_. **ScaleMode**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **ScaleMode** property uses the following settings.



|**Setting**|**Description**|
|:-----|:-----|
|0|Custom values used by one or more of the  **ScaleHeight**, **ScaleWidth**, **ScaleLeft**, and **ScaleTop** properties|
|1|(Default) Twips|
|2|Points|
|3|Pixels|
|4|Characters (horizontal = 120 twips per unit; vertical = 240 twips per unit)|
|5|Inches|
|6|Millimeters|
|7|Centimeters|
You can set the  **ScaleMode** property by using a macro or a[Visual Basic](set-properties-by-using-visual-basic.md)event procedure specified by a section's **OnPrint** property setting.

By using the related  **ScaleHeight**, **ScaleWidth**, **ScaleLeft**, and **ScaleTop** properties, you can create a custom coordinate system with both positive and negative coordinates. All four properties interact with the **ScaleMode** property in the following ways:


- Setting any other Scale property to any value automatically sets the  **ScaleMode** property to 0.
    
- Setting the  **ScaleMode** property to a number greater than 0 changes the **ScaleHeight** and **ScaleWidth** property settings to the new unit of measurement and sets the **ScaleLeft** and **ScaleTop** properties to 0. Also, the **CurrentX** and **CurrentY** property settings change to reflect the new coordinates of the current point.
    

## See also


#### Concepts


[Report Object](report-object-access.md)

