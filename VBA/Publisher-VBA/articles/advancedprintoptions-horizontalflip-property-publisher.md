---
title: AdvancedPrintOptions.HorizontalFlip Property (Publisher)
keywords: vbapb10.chm7077892
f1_keywords:
- vbapb10.chm7077892
ms.prod: publisher
api_name:
- Publisher.AdvancedPrintOptions.HorizontalFlip
ms.assetid: afb61c80-4706-8602-e78a-be35e2966c8c
ms.date: 06/08/2017
---


# AdvancedPrintOptions.HorizontalFlip Property (Publisher)

 **True** to print a horizontally mirrored image of the specified publication. The default is **False**. Read/write  **Boolean**.


## Syntax

 _expression_. **HorizontalFlip**

 _expression_A variable that represents an  **AdvancedPrintOptions** object.


## Remarks

This property is accessible only if the active printer is a PostScript printer. Returns a run-time error if a non-PostScript printer is specified. Use the  **[IsPostscriptPrinter](advancedprintoptions-ispostscriptprinter-property-publisher.md)** property of the **[AdvancedPrintOptions](advancedprintoptions-object-publisher.md)** object to determine if the specified printer is a PostScript printer.

This property is saved as an application setting and applied to future instances of Microsoft Publisher.

This property corresponds to the  **Flip horizontally** control on the **Page Settings** tab of the **Advanced Print Settings** dialog box.

This property is mostly used when printing to film on an imagesetter so that the image reads correctly when the emulsion side of the film is down (as when burning a press plate).


## Example

The following example determines if the active printer is a PostScript printer. If it is, the active publication is set to print as a horizontally mirrored and vertically mirrored, negative image of itself.


```vb
Sub PrepToPrintToFilmOnImagesetter() 
 
With ActiveDocument.AdvancedPrintOptions 
 If .IsPostscriptPrinter = True Then 
 .HorizontalFlip = True 
 .VerticalFlip = True 
 .NegativeImage = True 
 End If 
End With 
 
End Sub
```


## See also


#### Concepts


 [AdvancedPrintOptions Object](advancedprintoptions-object-publisher.md)

