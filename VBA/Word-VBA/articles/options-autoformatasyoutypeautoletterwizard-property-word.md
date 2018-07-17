---
title: Options.AutoFormatAsYouTypeAutoLetterWizard Property (Word)
keywords: vbawd10.chm162988336
f1_keywords:
- vbawd10.chm162988336
ms.prod: word
api_name:
- Word.Options.AutoFormatAsYouTypeAutoLetterWizard
ms.assetid: be49edd1-cb44-12d1-df43-ddaaddccef04
ms.date: 06/08/2017
---


# Options.AutoFormatAsYouTypeAutoLetterWizard Property (Word)

 **True** for Microsoft Word to automatically start the Letter Wizard when the user enters a letter salutation or closing. Read/write.


## Syntax

 _expression_ . **AutoFormatAsYouTypeAutoLetterWizard**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example sets Microsoft Word to automatically start the Letter Wizard when the user enters a letter salutation or closing.


```vb
Sub AutoLeterWizard() 
 Options.AutoFormatAsYouTypeAutoLetterWizard = True 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

