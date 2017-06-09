---
title: AdvancedPrintOptions.UseOnlyPublicationFonts Property (Publisher)
keywords: vbapb10.chm7077894
f1_keywords:
- vbapb10.chm7077894
ms.prod: publisher
api_name:
- Publisher.AdvancedPrintOptions.UseOnlyPublicationFonts
ms.assetid: f5973b32-37f3-8f65-1437-a465aa488ef4
ms.date: 06/08/2017
---


# AdvancedPrintOptions.UseOnlyPublicationFonts Property (Publisher)

Returns or sets a  **Boolean** that represents whether to only use publication fonts for printing the specified publication. **True** to print the specified publication using only fonts downloaded from your computer. Read/write. The default is **True**.


## Syntax

 _expression_. **UseOnlyPublicationFonts**

 _expression_A variable that represents an  **AdvancedPrintOptions** object.


### Return Value

Boolean


## Remarks

Publication fonts are fonts that are downloaded from your computer, as opposed to fonts residing at the printer or imagesetter.

Set this property to  **False** to enable the printer to print the specified publication using its resident fonts (stored in ROM, RAM, or on a hard disk drive) that have the same name as the fonts downloaded from your computer.


 **Note**  This may result in the printer substituting resident printer for fonts downloaded from your computer. This results in a slightly faster print time. However, if the resident fonts are not exactly identical to your computer fonts (even if they have the same name), this may cause your printed publication to look different than expected.

Setting this property to  **True** ensures that the fonts used to print the publication are the same ones used to create it.

This property corresponds to the  **Fonts** controls on the **Graphics and Fonts** tab of the **Advanced Print Settings** dialog box.


## Example

The following example tests to determine if the active publication will be printed using only publication fonts. If it will not, it is set to use only publication fonts.


```vb
Sub PrintWithPublicationFontsOnly() 
 With ActiveDocument.AdvancedPrintOptions 
 .UseOnlyPublicationFonts = True 
 End With 
End Sub
```


## See also


#### Concepts


 [AdvancedPrintOptions Object](advancedprintoptions-object-publisher.md)

