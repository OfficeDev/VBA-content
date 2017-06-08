---
title: AutoTextEntry Object (Word)
keywords: vbawd10.chm2358
f1_keywords:
- vbawd10.chm2358
ms.prod: word
api_name:
- Word.AutoTextEntry
ms.assetid: 37a2cf05-ae07-d411-9bd8-ab4726b303a9
ms.date: 06/08/2017
---


# AutoTextEntry Object (Word)

Represents a single AutoText entry. The  **AutoTextEntry** object is a member of the **AutoTextEntries** collection. The **[AutoTextEntries](autotextentries-object-word.md)** collection contains all the AutoText entries in the specified template. The entries are listed on the **AutoText** tab in the **AutoCorrect** dialog box.


## Remarks

Use  **[AutoTextEntries](autotextentries-item-method-word.md)** (index), where index is the AutoText entry name or index number, to return a single **AutoTextEntry** object. You must exactly match the spelling (but not necessarily the capitalization) of the name, as it is shown on the **AutoText** tab in the **AutoCorrect** dialog box. The following example sets the value of an existing AutoText entry named "cName."


```
NormalTemplate.AutoTextEntries("cName").Value = _ 
 "The Johnson Company"
```

The following example displays the name and value of the first AutoText entry in the template attached to the active document.




```vb
Set myTemplate = ActiveDocument.AttachedTemplate 
MsgBox "Name = " &; myTemplate.AutoTextEntries(1).Name &; vbCr _ 
 &; "Value " &; myTemplate.AutoTextEntries(1).Value
```

The following example inserts the global AutoText entry named "TheWorld" at the insertion point.




```
Selection.Collapse Direction:=wdCollapseEnd 
NormalTemplate.AutoTextEntries("TheWorld").Insert _ 
 Where:=Selection.Range
```

Use the  **[Add](autotextentries-add-method-word.md)** method to add an **AutoTextEntry** object to the **AutoTextEntries** collection. The following example adds an AutoText entry named "Blue" based on the text of the selection.




```
NormalTemplate.AutoTextEntries.Add Name:="Blue", _ 
 Range:=Selection.Range
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


