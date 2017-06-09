---
title: Printer.ItemLayout Property (Access)
keywords: vbaac10.chm12878
f1_keywords:
- vbaac10.chm12878
ms.prod: access
api_name:
- Access.Printer.ItemLayout
ms.assetid: 5e90c2fb-cc1a-48fb-d3c3-914c89737c74
ms.date: 06/08/2017
---


# Printer.ItemLayout Property (Access)

Returns or sets an  **[AcPrintItemLayout](acprintitemlayout-enumeration-access.md)** constant indicating whether the printer lays columns across, then down, or down, then across. Read/write.


## Syntax

 _expression_. **ItemLayout**

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

