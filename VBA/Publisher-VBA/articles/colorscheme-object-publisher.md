---
title: ColorScheme Object (Publisher)
keywords: vbapb10.chm2752511
f1_keywords:
- vbapb10.chm2752511
ms.prod: publisher
api_name:
- Publisher.ColorScheme
ms.assetid: b4e554ef-f043-c963-e175-b7d5ba95c636
ms.date: 06/08/2017
---


# ColorScheme Object (Publisher)

Represents a color scheme, which is a set of eight colors used for the different elements of a publication. Each color is represented by a  **[ColorFormat](colorformat-object-publisher.md)** object. The **ColorScheme** object is a member of the **[ColorSchemes](colorschemes-object-publisher.md)** collection. The **ColorSchemes** collection contains all the color schemes available to Microsoft Publisher.
 


## Example

Use the  **[ColorScheme](document-colorscheme-property-publisher.md)** property of a **[Document](document-object-publisher.md)** object to return the color scheme for the current publication. The following example sets the fill value of three shapes on the first page to the return value (in RGB format) of three of the eight **ColorScheme** colors.
 

 

```
Sub ReturnColorsAndApplyToShapes() 
 Dim lngAccent1 As Long 
 Dim lngAccent2 As Long 
 Dim lngAccent3 As Long 
 
 With ActiveDocument 
 With .ColorScheme 
 lngAccent1 = .Colors(pbSchemeColorAccent1).RGB 
 lngAccent2 = .Colors(pbSchemeColorAccent2).RGB 
 lngAccent3 = .Colors(pbSchemeColorAccent3).RGB 
 End With 
 With .Pages(1) 
 .Shapes(1).Fill.ForeColor.RGB = lngAccent1 
 .Shapes(2).Fill.ForeColor.RGB = lngAccent2 
 .Shapes(3).Fill.ForeColor.RGB = lngAccent3 
 End With 
 End With 
 
End Sub
```

Use the  **[Name](colorscheme-name-property-publisher.md)** property to return a color scheme name. The following example lists in a text box all the color schemes available to Publisher.
 

 



```
Sub ListColorShemes() 
 
 Dim clrScheme As ColorScheme 
 Dim strSchemes As String 
 
 For Each clrScheme In Application.ColorSchemes 
 strSchemes = strSchemes &amp; clrScheme.Name &amp; vbLf 
 Next 
 ActiveDocument.Pages(1).Shapes.AddTextbox( _ 
 Orientation:=pbTextOrientationHorizontal, _ 
 Left:=72, Top:=72, Width:=400, Height:=500).TextFrame _ 
 .TextRange.Text = strSchemes 
 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](colorscheme-application-property-publisher.md)|
|[Colors](colorscheme-colors-property-publisher.md)|
|[Name](colorscheme-name-property-publisher.md)|
|[Parent](colorscheme-parent-property-publisher.md)|

