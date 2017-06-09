---
title: TextColumns.EvenlySpaced Property (Word)
keywords: vbawd10.chm158531684
f1_keywords:
- vbawd10.chm158531684
ms.prod: word
api_name:
- Word.TextColumns.EvenlySpaced
ms.assetid: 9498889e-0f61-ddad-df6b-6defb11dc566
ms.date: 06/08/2017
---


# TextColumns.EvenlySpaced Property (Word)

 **True** if text columns are evenly spaced. Read/write **Long** .


## Syntax

 _expression_ . **EvenlySpaced**

 _expression_ A variable that represents a **[TextColumns](textcolumns-objectword.md)** object.


## Remarks

The  **EvenlySpaced** property can be **True** , **False** , or **wdUndefined** .

If you set the  **[Spacing](textcolumns-spacing-property-word.md)** or **[Width](textcolumns-width-property-word.md)** property of the **TextColumns** object, the **EvenlySpaced** property is automatically set to **True** . Also, setting the **EvenlySpaced** property may change the settings for the **Spacing** and **Width** properties of the **TextColumns** object.


## Example

This example topic sets columns in the active document to be evenly spaced.


```vb
Dim colTextColumns 
 
Set colTextColumns = ActiveDocument.PageSetup.TextColumns 
 
If colTextColumns.Count > 1 Then _ 
 colTextColumns.EvenlySpaced = True 
End If
```

This example returns the status of the  **Equal column width** option in the **Columns** dialog box ( **Format** menu).




```vb
Dim lngSpaced As Long 
 
lngSpaced = ActiveDocument.PageSetup.TextColumns.EvenlySpaced
```


## See also


#### Concepts


[TextColumns Collection Object](textcolumns-objectword.md)

