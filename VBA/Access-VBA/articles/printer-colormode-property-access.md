---
title: Printer.ColorMode Property (Access)
keywords: vbaac10.chm12857
f1_keywords:
- vbaac10.chm12857
ms.prod: access
api_name:
- Access.Printer.ColorMode
ms.assetid: 5c54604b-ee6a-2d6a-1a3e-3fea397a2fa0
ms.date: 06/08/2017
---


# Printer.ColorMode Property (Access)

Returns or sets an  **[AcPrintColor](acprintcolor-enumeration-access.md)** constant representing whether the specified printer should print output in color or monochrome. Read/write.


## Syntax

 _expression_. **ColorMode**

 _expression_ A variable that represents a **Printer** object.


## Example

The following example sets a variety of printer settings for the form specified in the  _strFormname_ argument of the procedure.


```vb
Sub SetPrinter(strFormname As String) 
 
 DoCmd.OpenForm FormName:=strFormname, view:=acDesign, _ 
 datamode:=acFormEdit, windowmode:=acHidden 
 
 With Forms(form1).Printer 
 
 .TopMargin = 1440 
 .BottomMargin = 1440 
 .LeftMargin = 1440 
 .RightMargin = 1440 
 
 .ColumnSpacing = 360 
 .RowSpacing = 360 
 
  .ColorMode = acPRCMColor 
 .DataOnly = False 
 .DefaultSize = False 
 .ItemSizeHeight = 2880 
 .ItemSizeWidth = 2880 
 .ItemLayout = acPRVerticalColumnLayout 
 .ItemsAcross = 6 
 
 .Copies = 1 
 .Orientation = acPRORLandscape 
 .Duplex = acPRDPVertical 
 .PaperBin = acPRBNAuto 
 .PaperSize = acPRPSLetter 
 .PrintQuality = acPRPQMedium 
 
 End With 
 
 DoCmd.Close objecttype:=acForm, objectname:=strFormname, _ 
 Save:=acSaveYes 
 
 
End Sub
```


## See also


#### Concepts


[Printer Object](printer-object-access.md)

