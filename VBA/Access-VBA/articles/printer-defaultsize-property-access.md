---
title: Printer.DefaultSize Property (Access)
keywords: vbaac10.chm12875
f1_keywords:
- vbaac10.chm12875
ms.prod: access
api_name:
- Access.Printer.DefaultSize
ms.assetid: b5dd3ce8-a5db-7562-5760-fc07c4409130
ms.date: 06/08/2017
---


# Printer.DefaultSize Property (Access)

 **True** if the size of the detail section in Design View is used for printing; otherwise, the values of the **[ItemSizeHeight](printer-itemsizeheight-property-access.md)** and **[ItemSizeWidth](printer-itemsizewidth-property-access.md)** properties are used. Read/write **Boolean**.


## Syntax

 _expression_. **DefaultSize**

 _expression_ A variable that represents a **Printer** object.


## Remarks

When this property is  **True**, the **ItemSizeHeight** and **ItemSizeWidth** properties are ignored.


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

