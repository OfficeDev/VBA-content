---
title: HeadingStyle Object (Word)
keywords: vbawd10.chm2443
f1_keywords:
- vbawd10.chm2443
ms.prod: word
api_name:
- Word.HeadingStyle
ms.assetid: d57e68ce-4c8b-0063-5077-82462451f336
ms.date: 06/08/2017
---


# HeadingStyle Object (Word)

Represents a style used to build a table of contents or figures. The  **HeadingStyle** object is a member of the **[HeadingStyles](headingstyles-object-word.md)** collection.


## Remarks

Use  **HeadingStyles** (Index), where Index is the index number, to return a single **HeadingStyle** object. The index number represents the position of the style in the **HeadingStyles** collection. The following example adds (at the beginning of the active document) a table of figures built from the Title style, and then displays the name of the first style in the **HeadingStyles** collection.


```vb
Set myTOF = ActiveDocument.TablesOfFigures.Add _ 
 (Range:=ActiveDocument.Range(0, 0), AddedStyles:="Title") 
MsgBox myTOF.HeadingStyles(1).Style
```

Use the  **Add** method to add a style to the **HeadingStyles** collection. The following example adds a table of contents at the beginning of the active document and then adds the Title style to the list of styles used to build a table of contents.




```vb
Set myToc = ActiveDocument.TablesOfContents.Add _ 
 (Range:=ActiveDocument.Range(0, 0), UseHeadingStyles:=True, _ 
 LowerHeadingLevel:=3, UpperHeadingLevel:=1) 
myToc.HeadingStyles.Add Style:="Title", Level:=2
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

