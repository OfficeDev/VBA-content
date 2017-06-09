---
title: HTMLDivision Object (Word)
keywords: vbawd10.chm2535
f1_keywords:
- vbawd10.chm2535
ms.prod: word
api_name:
- Word.HTMLDivision
ms.assetid: a38918ed-61aa-3fd1-3522-d077f1ff312f
ms.date: 06/08/2017
---


# HTMLDivision Object (Word)

Represents a single HTML DIV element within a Web document. The  **HTMLDivision** object is a member of the **HTMLDivisions** collection.


## Remarks

Use  **HTMLDivisions** (Index), where Index refers to the HTML division in the document, to return a single **HTMLDivision** object. Use the **Borders** property to format border properties for an HTML division. This example formats three nested divisions in the active document. This example assumes that the active document is an HTML document with at least three divisions.


```vb
Sub FormatHTMLDivisions() 
 With ActiveDocument.HTMLDivisions(1) 
 With .Borders(wdBorderLeft) 
 .Color = wdColorRed 
 .LineStyle = wdLineStyleSingle 
 End With 
 With .Borders(wdBorderTop) 
 .Color = wdColorRed 
 .LineStyle = wdLineStyleSingle 
 End With 
 With .HTMLDivisions(1) 
 .LeftIndent = InchesToPoints(1) 
 .RightIndent = InchesToPoints(1) 
 With .Borders(wdBorderRight) 
 .Color = wdColorBlue 
 .LineStyle = wdLineStyleDouble 
 End With 
 End With 
 With .Borders(wdBorderBottom) 
 .Color = wdColorBlue 
 .LineStyle = wdLineStyleDouble 
 End With 
 With .HTMLDivisions(1) 
 .LeftIndent = InchesToPoints(1) 
 .RightIndent = InchesToPoints(1) 
 With .Borders(wdBorderLeft) 
 .Color = wdColorBlack 
 .LineStyle = wdLineStyleDot 
 End With 
 With .Borders(wdBorderTop) 
 .Color = wdColorBlack 
 .LineStyle = wdLineStyleDot 
 End With 
 End With 
 End With 
 End With 
 
End Sub
```

HTML divisions can be nested within multiple HTML divisions. Use the  **HTMLDivisionParent** method to access a parent HTML division of the current HTML division. This example formats the borders for two HTML divisions in the active document. This example assumes that the active document is an HTML document with at least two divisions.




```vb
Sub FormatHTMLDivisions() 
 With ActiveDocument.HTMLDivisions(1) 
 With .HTMLDivisions(1) 
 .LeftIndent = InchesToPoints(1) 
 .RightIndent = InchesToPoints(1) 
 With .Borders(wdBorderLeft) 
 .Color = wdColorBlue 
 .LineStyle = wdLineStyleDouble 
 End With 
 With .Borders(wdBorderRight) 
 .Color = wdColorBlue 
 .LineStyle = wdLineStyleDouble 
 End With 
 With .HTMLDivisionParent 
 .LeftIndent = InchesToPoints(1) 
 .RightIndent = InchesToPoints(1) 
 With .Borders(wdBorderTop) 
 .Color = wdColorBlack 
 .LineStyle = wdLineStyleDot 
 End With 
 With .Borders(wdBorderBottom) 
 .Color = wdColorBlack 
 .LineStyle = wdLineStyleDot 
 End With 
 End With 
 End With 
 End With 
End Sub
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


