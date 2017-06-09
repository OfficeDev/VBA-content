---
title: Page Object (Word)
keywords: vbawd10.chm169
f1_keywords:
- vbawd10.chm169
ms.prod: word
api_name:
- Word.Page
ms.assetid: 3a3d480a-3876-515f-d13f-7ec23818245f
ms.date: 06/08/2017
---


# Page Object (Word)

Represents a page in a document. Use the  **Page** object and the related methods and properties for programmatically defining page layout in a document.


## Remarks

Use the  **Item** method to access a specific page in a document. The following example accesses the first page in the active document.


```
Dim objPage As Page 
 
Set objPage = ActiveDocument.ActiveWindow _ 
 .Panes(1).Pages.Item(1)
```

To access the page number, use the  **Information** property of a **Range** or **Selection** object, or the **PageIndex** property of a **Break** object that belongs to the **Breaks** collection of the specified **Page** object.

The  **Top** and **Left** properties of the **Page** object always return 0 (zero) indicating the upper left corner of the page. The **Height** and **Width** properties return the height and width in points (72 points = 1 inch) of the paper size specified in the Page Setup dialog or through the **PageSetup** object. For example, for an 8-1/2 by 11 inch page in portrait mode, the **Height** property returns 792 and the **Width** property returns 612. All four of these properties are read-only.


## Methods



|**Name**|
|:-----|
|[SaveAsPNG](http://msdn.microsoft.com/library/f734988c-2cea-2a51-66b5-d3e7c6c30d56%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](page-application-property-word.md)|
|[Breaks](page-breaks-property-word.md)|
|[Creator](page-creator-property-word.md)|
|[EnhMetaFileBits](page-enhmetafilebits-property-word.md)|
|[Height](page-height-property-word.md)|
|[Left](page-left-property-word.md)|
|[Parent](page-parent-property-word.md)|
|[Rectangles](page-rectangles-property-word.md)|
|[Top](page-top-property-word.md)|
|[Width](page-width-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
