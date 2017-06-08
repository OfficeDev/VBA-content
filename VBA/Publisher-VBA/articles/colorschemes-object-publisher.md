---
title: ColorSchemes Object (Publisher)
keywords: vbapb10.chm2818047
f1_keywords:
- vbapb10.chm2818047
ms.prod: publisher
api_name:
- Publisher.ColorSchemes
ms.assetid: f5002de1-5e91-fc92-eedb-0e13dce57802
ms.date: 06/08/2017
---


# ColorSchemes Object (Publisher)

A collection of all the  **[ColorScheme](colorscheme-object-publisher.md)** objects in Microsoft Publisher. Each **ColorScheme** object represents a color scheme, which is a set of colors that are used in a publication.
 


## Example

Use the  **[Count](colorschemes-count-property-publisher.md)** property to return the number of color schemes available to Publisher. The following example displays the number of color schemes.
 

 

```
Sub CountColorSchemes() 
 MsgBox Application.ColorSchemes.Count 
End Sub
```

Use the  **[Item](colorschemes-item-property-publisher.md)** property to return a specific color scheme from the **ColorSchemes** collection. The ** _Index_** argument of the **Item** property can be the number or name of the color scheme, or a **PbColorScheme** constant. The follow example sets the color scheme of the active publication to Wildflower.
 

 



```
Sub SetColorScheme() 
 ActiveDocument.ColorScheme _ 
 = ColorSchemes.Item(pbColorSchemeWildflower) 
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
|[Application](colorschemes-application-property-publisher.md)|
|[Count](colorschemes-count-property-publisher.md)|
|[Item](colorschemes-item-property-publisher.md)|
|[Parent](colorschemes-parent-property-publisher.md)|

