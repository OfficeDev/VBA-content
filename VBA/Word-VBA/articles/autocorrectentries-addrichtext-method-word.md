---
title: AutoCorrectEntries.AddRichText Method (Word)
keywords: vbawd10.chm155713638
f1_keywords:
- vbawd10.chm155713638
ms.prod: word
api_name:
- Word.AutoCorrectEntries.AddRichText
ms.assetid: e03f37ca-1011-825f-5a79-29a23f2371f0
ms.date: 06/08/2017
---


# AutoCorrectEntries.AddRichText Method (Word)

Creates a formatted AutoCorrect entry, preserving all text attributes of the specified range. Returns an  **AutoCorrectEntry** object.


## Syntax

 _expression_ . **AddRichText**( **_Name_** , **_Range_** )

 _expression_ Required. A variable that represents an **[AutoCorrectEntries](autocorrectentries-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The text to replace automatically with Range.|
| _Range_|Required| **Range**|The formatted text that Word will insert automatically whenever Name is typed.|

### Return Value

AutoCorrectEntry


## Remarks

The  **RichText** property for entries added by using this method returns **True** . If **AddRichText** isn't used, inserted **AutoCorrect** entries conform to the current style.


## Example

This example stores the selected text as a formatted AutoCorrect entry that will be inserted automatically whenever "NewText" is typed.


```vb
If Selection.Type = wdSelectionNormal Then 
 AutoCorrect.Entries.AddRichText "NewText", Selection.Range 
Else 
 MsgBox "You need to select some text." 
End If
```

This example stores the third word in the active document as a formatted AutoCorrect entry that will be inserted automatically whenever "NewText" is typed.




```
AutoCorrect.Entries.AddRichText "NewText", ActiveDocument.Words(3)
```


## See also


#### Concepts


[AutoCorrectEntries Collection Object](autocorrectentries-object-word.md)

