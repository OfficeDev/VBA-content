---
title: Options.CheckSpellingAsYouType Property (Word)
keywords: vbawd10.chm162988308
f1_keywords:
- vbawd10.chm162988308
ms.prod: word
api_name:
- Word.Options.CheckSpellingAsYouType
ms.assetid: 8e4b55af-8fc6-2c99-ebfb-f008657d0da6
ms.date: 06/08/2017
---


# Options.CheckSpellingAsYouType Property (Word)

 **True** if Microsoft Word checks spelling and marks errors automatically as you type. Read/write **Boolean** .


## Syntax

 _expression_ . **CheckSpellingAsYouType**

 _expression_ A variable that represents a **[Options](options-object-word.md)** object.


## Remarks

This property marks spelling errors, but to see them on the screen, you must set the  **[ShowSpellingErrors](document-showspellingerrors-property-word.md)** property to **True** .


## Example

This example turns off automatic checking of spelling in Word.


```vb
Options.CheckSpellingAsYouType = False
```

This example sets Word to check for spelling errors as you type and to display any errors found in the active document.




```vb
Options.CheckSpellingAsYouType = True 
ActiveDocument.ShowSpellingErrors = True
```

This example returns the status of the  **Check spelling as you type** option on the **Spelling &; Grammar** tab in the **Options** dialog box ( **Tools** menu).




```vb
Dim blnCheck As Boolean 
 
blnCheck = Options.CheckSpellingAsYouType
```


## See also


#### Concepts


[Options Object](options-object-word.md)

