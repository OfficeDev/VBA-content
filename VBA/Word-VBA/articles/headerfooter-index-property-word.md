---
title: HeaderFooter.Index Property (Word)
ms.prod: word
api_name:
- Word.HeaderFooter.Index
ms.assetid: 5281c150-1a61-670f-6b1f-37c43b717126
ms.date: 06/08/2017
---


# HeaderFooter.Index Property (Word)

Returns a  **WdHeaderFooterIndex** that represents the specified header or footer in a document or section. Read-only.


## Syntax

 _expression_ . **Index**

 _expression_ Required. A variable that represents a **[HeaderFooter](headerfooter-object-word.md)** object.


## Example

This example adds a shape to the first page header in the active document if the specified variable references the first page header.


```vb
Sub ChangeFirstPageFooter() 
 Dim hdrFirstPage As HeaderFooter 
 
 Set hdrFirstPage = ActiveDocument.Sections(1).Headers(wdHeaderFooterFirstPage) 
 
 If hdrFirstPage.Index = wdHeaderFooterFirstPage Then 
 With hdrFirstPage.Shapes.AddShape(Type:=msoShapeHeart, _ 
 Left:=36, Top:=36, Width:=36, Height:=36) 
 .Fill.ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=0) 
 End With 
 End If 
 
End Sub
```


## See also


#### Concepts


[HeaderFooter Object](headerfooter-object-word.md)

