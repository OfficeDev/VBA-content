---
title: HeaderFooter Object (Word)
ms.prod: word
api_name:
- Word.HeaderFooter
ms.assetid: 3f2f926a-9220-5536-80ed-af63d2feb016
ms.date: 06/08/2017
---


# HeaderFooter Object (Word)

Represents a single header or footer. The  **HeaderFooter** object is a member of the **[HeadersFooters](headersfooters-object-word.md)** collection. The **HeadersFooters** collection includes all headers and footers in the specified document section.


## Remarks

Use  **Headers** (Index) or **Footers** (Index), where index is one of the **WdHeaderFooterIndex** constants ( **wdHeaderFooterEvenPages**, **wdHeaderFooterFirstPage**, or **wdHeaderFooterPrimary** ), to return a single **HeaderFooter** object. The following example changes the text of both the primary header and the primary footer in the first section of the active document.


```
With ActiveDocument.Sections(1) 
 .Headers(wdHeaderFooterPrimary).Range.Text = "Header text" 
 .Footers(wdHeaderFooterPrimary).Range.Text = "Footer text" 
End With
```

You can also return a single  **HeaderFooter** object by using the **HeaderFooter** property with a **Selection** object.


 **Note**  You cannot add  **HeaderFooter** objects to the **[HeadersFooters](headersfooters-object-word.md)** collection.

Use the  **DifferentFirstPageHeaderFooter** property with the **PageSetup** object to specify a different first page. The following example inserts text into the first page footer in the active document.




```
With ActiveDocument 
 .PageSetup.DifferentFirstPageHeaderFooter = True 
 .Sections(1).Footers(wdHeaderFooterFirstPage) _ 
 .Range.InsertBefore _ 
 "Written by Joe Smith" 
End With
```

Use the  **OddAndEvenPagesHeaderFooter** property with the **PageSetup** object to specify different odd and even page headers and footers. If the **OddAndEvenPagesHeaderFooter** property is **True**, you can return an odd header or footer by using **wdHeaderFooterPrimary**, and you can return an even header or footer by using **wdHeaderFooterEvenPages**.

Use the  **Add** method with the **PageNumbers** object to add a page number to a header or footer. The following example adds page numbers to the primary footer in the first section of the active document.




```
With ActiveDocument.Sections(1) 
 .Footers(wdHeaderFooterPrimary).PageNumbers.Add 
End With
```


## Properties



|**Name**|
|:-----|
|[Application](headerfooter-application-property-word.md)|
|[Creator](headerfooter-creator-property-word.md)|
|[Exists](headerfooter-exists-property-word.md)|
|[Index](headerfooter-index-property-word.md)|
|[IsHeader](headerfooter-isheader-property-word.md)|
|[LinkToPrevious](headerfooter-linktoprevious-property-word.md)|
|[PageNumbers](headerfooter-pagenumbers-property-word.md)|
|[Parent](headerfooter-parent-property-word.md)|
|[Range](headerfooter-range-property-word.md)|
|[Shapes](headerfooter-shapes-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
