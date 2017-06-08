---
title: PageSetup Object (Word)
keywords: vbawd10.chm2417
f1_keywords:
- vbawd10.chm2417
ms.prod: word
api_name:
- Word.PageSetup
ms.assetid: 1879d601-80ad-4fc0-1a87-92e999b59f88
ms.date: 06/08/2017
---


# PageSetup Object (Word)

Represents the page setup description. The  **PageSetup** object contains all the page setup attributes of a document (such as left margin, bottom margin, and paper size) as properties.


## Remarks

Use the  **PageSetup** property to return the **PageSetup** object. The following example sets the first section in the active document to landscape orientation and then prints the document.


```vb
ActiveDocument.Sections(1).PageSetup.Orientation = _ 
 wdOrientLandscape 
ActiveDocument.PrintOut
```

The following example sets all the margins for the document named "Sales.doc."




```vb
With Documents("Sales.doc").PageSetup 
 .LeftMargin = InchesToPoints(0.75) 
 .RightMargin = InchesToPoints(0.75) 
 .TopMargin = InchesToPoints(1.5) 
 .BottomMargin = InchesToPoints(1) 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


