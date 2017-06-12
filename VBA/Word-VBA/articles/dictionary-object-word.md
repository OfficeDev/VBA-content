---
title: Dictionary Object (Word)
keywords: vbawd10.chm2477
f1_keywords:
- vbawd10.chm2477
ms.prod: word
api_name:
- Word.Dictionary
ms.assetid: 1946d60c-2abd-9ca9-8d0b-7068e9173bb3
ms.date: 06/08/2017
---


# Dictionary Object (Word)

Represents a dictionary.  **Dictionary** objects that represent custom dictionaries are members of the **[Dictionaries](dictionaries-object-word.md)** collection. Other dictionary objects are returned by properties of the **[Languages](languages-object-word.md)** collection; these include the **[ActiveSpellingDictionary](language-activespellingdictionary-property-word.md)** , **[ActiveGrammarDictionary](language-activegrammardictionary-property-word.md)** , **[ActiveThesaurusDictionary](language-activethesaurusdictionary-property-word.md)** , and **[ActiveHyphenationDictionary](language-activehyphenationdictionary-property-word.md)** properties.


## Remarks

Use  **[CustomDictionaries](application-customdictionaries-property-word.md)** (Index), where Index is an index number or the string name for the dictionary, to return a single **Dictionary** object that represents a custom dictionary. The following example returns the first dictionary in the collection.


```
CustomDictionaries(1)
```

The following example returns the dictionary named "MyDictionary."




```
CustomDictionaries("MyDictionary")
```

Use the  **[ActiveCustomDictionary](dictionaries-activecustomdictionary-property-word.md)** property to set the custom spelling dictionary in the collection to which new words are added. If you try to set this property to a dictionary that's not a custom dictionary, an error occurs.

Use the  **[Add](dictionaries-add-method-word.md)** method to add a new dictionary to the collection of active custom dictionaries. If there is no file with the name specified by FileName, Word creates it. The following example adds "MyCustom.dic" to the collection of custom dictionaries.




```
CustomDictionaries.Add FileName:="MyCustom.dic"
```

Remarks

Use the  **[Name](dictionary-name-property-word.md)** and **[Path](dictionary-path-property-word.md)** properties to locate any of the dictionaries. The following example displays a message box that contains the full path for each dictionary.




```vb
For Each d in CustomDictionaries 
 Msgbox d.Path &; Application.PathSeparator &; d.Name 
Next d
```

Use the  **[LanguageSpecific](dictionary-languagespecific-property-word.md)** property to determine whether the specified custom dictionary can have a specific language assigned to it with the **[LanguageID](dictionary-languageid-property-word.md)** property. If the dictionary is language specific, it will verify only text that's formatted for the specified language.

For each language for which proofing tools are installed, you can use the  **ActiveGrammarDictionary** , **ActiveHyphenationDictionary** , **ActiveSpellingDictionary** , and **ActiveThesaurusDictionary** properties to return the corresponding **Dictionary** objects. The following example returns the full path for the active spelling dictionary used in the U.S. English version of Word.




```vb
Set myspell = Languages(wdEnglishUS).ActiveSpellingDictionary 
MsgBox mySpell.Path &; Application.PathSeparator &; mySpell.Name
```

The  **[ReadOnly](dictionary-readonly-property-word.md)** property returns **True** for .lex files (built-in proofing dictionaries) and **False** for .dic files (custom spelling dictionaries).


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


