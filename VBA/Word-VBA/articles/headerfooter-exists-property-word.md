---
title: HeaderFooter.Exists Property (Word)
keywords: vbawd10.chm159711236
f1_keywords:
- vbawd10.chm159711236
ms.prod: word
api_name:
- Word.HeaderFooter.Exists
ms.assetid: 84ce3ac9-a4be-f99a-eb4b-1a145373659f
ms.date: 06/08/2017
---


# HeaderFooter.Exists Property (Word)

 **True** if the specified **HeaderFooter** object exists. Read/write **Boolean** .


## Syntax

 _expression_ . **Exists**

 _expression_ A variable that represents a **[HeaderFooter](headerfooter-object-word.md)** object.


## Remarks

The primary header and footer exist in all new documents by default. Use this method to determine whether a first-page or odd-page header or footer exists. You can also use the  **[DifferentFirstPageHeaderFooter](pagesetup-differentfirstpageheaderfooter-property-word.md)** or **[OddAndEvenPagesHeaderFooter](pagesetup-oddandevenpagesheaderfooter-property-word.md)** property to return or set the number of headers and footers in the specified document or section.


## Example

If a first-page header exists in section one, this example sets the text for the header.


```vb
Dim secTemp As Section 
 
Set secTemp = ActiveDocument.Sections(1) 
If secTemp.Headers(wdHeaderFooterFirstPage).Exists = True Then 
 secTemp.Headers(wdHeaderFooterFirstPage).Range.Text = _ 
 "First Page" 
End If
```


## See also


#### Concepts


[HeaderFooter Object](headerfooter-object-word.md)

