---
title: HeadersFooters Object (Word)
ms.prod: word
ms.assetid: 41dbbaa7-f139-3d3c-54d4-03a57ab8417a
ms.date: 06/08/2017
---


# HeadersFooters Object (Word)

A collection of  **[HeaderFooter](headerfooter-object-word.md)** objects that represent the headers or footers in the specified section of a document.


## Remarks

Use the  **Headers** or **Footers** property to return the **HeadersFooters** collection. The following example displays the text from the primary footer in the first section of the active document.


```vb
With ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary) 
 If .Range.Text <> vbCr Then 
 MsgBox .Range.Text 
 Else 
 MsgBox "Footer is empty" 
 End If 
End With
```


 **Note**  You cannot add  **HeaderFooter** objects to the **HeadersFooters** collection.

Use  **Headers** (Index) or **Footers** (Index), where index is one of the **WdHeaderFooterIndex** constants ( **wdHeaderFooterEvenPages** , **wdHeaderFooterFirstPage** , or **wdHeaderFooterPrimary** ), to return a single **HeaderFooter** object. The following example changes the text of both the primary header and the primary footer the first section of the active document.




```vb
With ActiveDocument.Sections(1) 
 .Headers(wdHeaderFooterPrimary).Range.Text = "Header text" 
 .Footers(wdHeaderFooterPrimary).Range.Text = "Footer text" 
End With
```

You can also return a single  **HeaderFooter** object by using the **HeaderFooter** property with a **Selection** object.

Use the  **DifferentFirstPageHeaderFooter** property with the **PageSetup** object to specify a different first page. The following example inserts text into the first page footer in the active document.




```vb
With ActiveDocument 
 .PageSetup.DifferentFirstPageHeaderFooter = True 
 .Sections(1).Footers(wdHeaderFooterFirstPage) _ 
 .Range.InsertBefore _ 
 "Written by Kate Edson" 
End With
```

Use the  **OddAndEvenPagesHeaderFooter** property with the **PageSetup** object to specify different odd and even page headers and footers. If the **OddAndEvenPagesHeaderFooter** property is **True** , you can return an odd header or footer by using **wdHeaderFooterPrimary** , and you can return an even header or footer by using **wdHeaderFooterEvenPages** .

Use the  **Add** method with the **PageNumbers** object to add a page number to a header or footer. The following example adds page numbers to the first page footer in the first section in the active document.




```vb
With ActiveDocument.Sections(1) 
 .PageSetup.DifferentFirstPageHeaderFooter = True 
 .Footers(wdHeaderFooterPrimary).PageNumbers.Add _ 
 FirstPage:=True 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


