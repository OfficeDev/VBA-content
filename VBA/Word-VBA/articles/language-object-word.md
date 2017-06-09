---
title: Language Object (Word)
keywords: vbawd10.chm2413
f1_keywords:
- vbawd10.chm2413
ms.prod: word
api_name:
- Word.Language
ms.assetid: 0acc4a42-b4c2-a415-0e38-a049b085dc86
ms.date: 06/08/2017
---


# Language Object (Word)

Represents a language used for proofing or formatting in Microsoft Word. The  **Language** object is a member of the **Languages** collection.


## Remarks

Use  **Languages** (Index) to return a single **Language** object, where Index can be the value of the **Name** property, the value of the **NameLocal** property, one of the **WdLanguageID** constants, or one of the **MsoLanguageID** constants. (For the list of valid **WdLanguageID** or **MsoLanguageID** constants, see the Object Browser in the Visual Basic Editor.)

The  **Name** property returns the name of a language, whereas the **NameLocal** property returns the name of a language in the language of the user. The following example returns the string "Italiano" for **Name** and "Italian (Standard)" for **NameLocal** when it is run in the U.S. English version of Word.




```vb
Sub ShowItalianNames() 
 Msgbox Languages(wdItalian).Name 
 Msgbox Languages(wdItalian).NameLocal 
End Sub
```

For each language for which proofing tools are installed, you can use the  **ActiveGrammarDictionary** , **ActiveHyphenationDictionary** , **ActiveSpellingDictionary** , and **ActiveThesaurusDictionary** properties to return the corresponding **Dictionary** object. The following example returns the full path for the active spelling dictionary used in the U.S. English version of Word.




```vb
Sub ShowDictionaryPath 
 Set myspell = Languages(wdEnglishUS).ActiveSpellingDictionary 
 MsgBox mySpell.Path &; Application.PathSeparator &; mySpell.Name 
End Sub
```

The writing style is the set of rules used by the grammar checker. The  **WritingStyleList** property returns an array of strings that represent the available writing styles for the specified language. The following example returns the list of writing styles for U.S. English.




```vb
Sub ListWritingStyles() 
 WrStyles = Languages(wdEnglishUS).WritingStyleList 
 For i = 1 To UBound(WrStyles) 
 MsgBox WrStyles(i) 
 Next i 
End Sub
```

Use the  **DefaultWritingStyle** property to set the default writing style you want Word to use.




```
Languages(wdEnglishUS).DefaultWritingStyle = "Casual"
```

You can override the default writing style with the  **ActiveWritingStyle** property. This property is applied to a specified document for text marked in a specified language. The following example sets the writing style to be used for checking U.S. English, French, and German in the active document.




```vb
Sub SetWritingStyle() 
 With ActiveDocument 
 .ActiveWritingStyle(wdEnglishUS) = "Technical" 
 .ActiveWritingStyle(wdFrench) = "Commercial" 
 .ActiveWritingStyle(wdGerman) = "Technisch/Wiss" 
 End With 
End Sub
```

If you mark text as  **wdNoProofing** , Word skips the marked text when running a spelling or grammar check.


 **Note**  You must have the proofing tools installed for each language you intend to check. For more information on working in other languages, see [Language-specific information](http://msdn.microsoft.com/library/b27a2d10-8a15-7b36-b329-34d55ada9f37%28Office.15%29.aspx).


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


