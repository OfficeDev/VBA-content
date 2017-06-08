---
title: AdvancedPrintOptions.IsPostscriptPrinter Property (Publisher)
keywords: vbapb10.chm7077921
f1_keywords:
- vbapb10.chm7077921
ms.prod: publisher
api_name:
- Publisher.AdvancedPrintOptions.IsPostscriptPrinter
ms.assetid: 69c31e55-2781-38fa-7c4a-c5bc1b49972a
ms.date: 06/08/2017
---


# AdvancedPrintOptions.IsPostscriptPrinter Property (Publisher)

Returns  **True** if the active printer is a PostScript printer. Read-only **Boolean**.


## Syntax

 _expression_. **IsPostscriptPrinter**

 _expression_A variable that represents an  **AdvancedPrintOptions** object.


### Return Value

Boolean


## Remarks

The following properties of the  **[AdvancedPrintOptions](advancedprintoptions-object-publisher.md)** object are only accessible if the active printer is a Postscript printer: **[HorizontalFlip](advancedprintoptions-horizontalflip-property-publisher.md)**,  **[VerticalFlip](advancedprintoptions-verticalflip-property-publisher.md)**, and  **[NegativeImage](advancedprintoptions-negativeimage-property-publisher.md)**.

Use the  **[IsActivePrinter](printer-isactiveprinter-property-publisher.md)** property to specify the active printer for a publication.


## Example

The following example determines if the active printer is a PostScript printer. If it is, the active publication is set to print as a horizontally and vertically mirrored, negative image of itself.


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

