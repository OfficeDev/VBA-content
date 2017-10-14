---
title: ListLevels Object (Word)
ms.prod: word
ms.assetid: 9165c008-c066-8d3e-9254-d9e0ab2ec091
ms.date: 06/08/2017
---


# ListLevels Object (Word)

A collection of  **ListLevel** objects that represents all the list levels of a list template, either the only level for a bulleted or numbered list or one of the nine levels of an outline numbered list.


## Remarks

Use the  **ListLevels** property to return the **ListLevels** collection. The following example sets the variable _mytemp_ to the first list template in the active document and then modifies each level to use lowercase letters for its number style.


```vb
Set mytemp = ActiveDocument.ListTemplates(1) 
For Each lev In mytemp.ListLevels 
 lev.NumberStyle = wdListNumberStyleLowercaseLetter 
Next lev
```

Use  **ListLevels** (Index), where Index is a number from 1 through 9, to return a single **[ListLevel](listlevel-object-word.md)** object. The following example sets list level one of list template one in the active document to start at four.




```vb
ActiveDocument.ListTemplates(1).ListLevels(1).StartAt = 4
```


 **Note**  You cannot add new levels to a list template.

To apply a list level, first identify the range or list, and then use the  **ApplyListTemplate** method. Each tab at the beginning of the paragraph is translated into a list level. For example, a paragraph that begins with three tabs will become a level-three list paragraph after the **ApplyListTemplate** method is used.


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


