---
title: ListLevel.NumberStyle Property (Word)
keywords: vbawd10.chm160235524
f1_keywords:
- vbawd10.chm160235524
ms.prod: word
api_name:
- Word.ListLevel.NumberStyle
ms.assetid: 1118eb25-3b57-3a9b-6323-ba8233636f3b
ms.date: 06/08/2017
---


# ListLevel.NumberStyle Property (Word)

Returns or sets the number style for the  **ListLevel** object. Read/write **WdListNumberStyle** .


## Syntax

 _expression_ . **NumberStyle**

 _expression_ Required. A variable that represents a **[ListLevel](listlevel-object-word.md)** object.


## Remarks

Some of the  **WdListNumberStyle** constants may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed.


## Example

This example creates an alternating number style for the third outline-numbered list template.


```vb
Set myTemp = ListGalleries(wdOutlineNumberGallery).ListTemplates(3) 
For i = 1 to 9 
 If i Mod 2 = 0 Then 
 myTemp.ListLevels(i).NumberStyle = _ 
 wdListNumberStyleUppercaseRoman 
 Else 
 myTemp.ListLevels(i).NumberStyle = _ 
 wdListNumberStyleLowercaseRoman 
 End If 
Next i
```

This example changes the number style to uppercase letters for every outline-numbered list in the active document.




```vb
For Each lt In ActiveDocument.ListTemplates 
 For Each ll In lt.listlevels 
 ll.NumberStyle = wdListNumberStyleUppercaseLetter 
 Next ll 
Next lt
```


## See also


#### Concepts


[ListLevel Object](listlevel-object-word.md)

