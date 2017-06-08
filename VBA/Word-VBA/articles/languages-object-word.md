---
title: Languages Object (Word)
ms.prod: word
ms.assetid: e3b1d3f3-de1b-d2fe-962f-5a589842d1b0
ms.date: 06/08/2017
---


# Languages Object (Word)

A collection of  **Language** objects that represent languages used for proofing or formatting in Word.


## Remarks

Use the  **Languages** property to return the **Languages** collection. The following example displays the localized name for each language.


```vb
For Each la In Languages 
 Msgbox la.NameLocal 
Next la
```

Use  **Languages** (index) to return a single **[Language](language-object-word.md)** object, where index can be the value of the **Name** property, the value of the **NameLocal** property, one of the **[WdLanguageID](wdlanguageid-enumeration-word.md)** constants, or one of the **MsoLanguageID** constants. (For the list of valid **MsoLanguageID** constants, see the Object Browser in the Visual Basic Editor.)

The  **Count** property returns the number of languages for which you can mark text (languages for which proofing tools are available). To check proofing, you must install the appropriate tools for each language you intend to check. You need both a .dll file and an .lex file for each of the following: the thesaurus, spelling checker, grammar checker, and hyphenation tools.

If you mark text as  **wdNoProofing** , Word skips the marked text when running a spelling or grammar check. To mark text for a specified language or for no proofing, use the **Set Language** command.


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


