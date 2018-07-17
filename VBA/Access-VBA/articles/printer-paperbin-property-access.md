---
title: Printer.PaperBin Property (Access)
keywords: vbaac10.chm12863
f1_keywords:
- vbaac10.chm12863
ms.prod: access
api_name:
- Access.Printer.PaperBin
ms.assetid: d3e33714-0aa5-aa9e-2b66-86afca3b38ee
ms.date: 06/08/2017
---


# Printer.PaperBin Property (Access)

Returns or sets an  **[AcPrintPaperBin](acprintpaperbin-enumeration-access.md)** constant indicating which paper bin the specified printer should use. Read/write.


## Syntax

 _expression_. **PaperBin**

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

