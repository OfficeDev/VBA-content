---
title: Dictionaries.Add Method (Word)
keywords: vbawd10.chm162267237
f1_keywords:
- vbawd10.chm162267237
ms.prod: word
api_name:
- Word.Dictionaries.Add
ms.assetid: aacd7041-e34f-b6e4-d895-925dad575d40
ms.date: 06/08/2017
---


# Dictionaries.Add Method (Word)

Returns a  **Dictionary** object that represents a new custom spelling or conversion dictionary added to the collection of active custom spelling or conversion dictionaries.


## Syntax

 _expression_ . **Add**( **_FileName_** )

 _expression_ Required. A variable that represents a **[Dictionaries](dictionaries-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The string name of the dictionary file. If no path is specified in the string, the proofing tools path is used.|

### Return Value

Dictionary


## Remarks

If a file with the name specified by the FileName parameter doesn't exist, Microsoft Word creates one.

The  **Dictionaries** collection includes only the active custom spelling dictionaries. **Dictionary** objects that are derived from the **Languages** collection don't have an **Add** method. These include the **Dictionary** objects returned by the **ActiveSpellingDictionary** , **ActiveGrammarDictionary** , **ActiveThesaurusDictionary** , and **ActiveHyphenationDictionary** properties.

Use the  **HangulHanjaDictionaries** property to return the collection of custom conversion dictionaries. The **HangulHanjaConversionDictionaries** collection includes only the active custom conversion dictionaries.


## Example

This example removes all dictionaries from the list of custom spelling dictionaries and creates a new custom dictionary file. The new dictionary is assigned to be the active custom dictionary, to which new words are automatically added.


```vb
With CustomDictionaries 
 .ClearAll 
 .Add FileName:="c:\My Documents\MyCustom.dic" 
 .ActiveCustomDictionary = CustomDictionaries(1) 
End With
```

This example creates a new custom dictionary and assigns it to a variable. The new custom dictionary is then set to be used for text that's marked as French Canadian. Note that to run a spelling check for another language, you must have installed the proofing tools for that language.




```vb
Sub FrCanDic() 
 Dim dicFrenchCan As Dictionary 
 
 Set dicFrenchCan = CustomDictionaries.Add(FileName:="FrenchCanadian.dic") 
 With dicFrenchCan 
 .LanguageSpecific = True 
 .LanguageID = wdFrenchCanadian 
 End With 
End Sub
```

This example removes all dictionaries from the list of custom conversion dictionaries and creates a new custom dictionary file. The new dictionary is assigned to be the active custom dictionary, to which new words are automatically added.




```vb
With HangulHanjaDictionaries 
 .ClearAll 
 .Add FileName:="C:\My Documents\MyCustom.hhd" 
 .ActiveCustomDictionary = CustomDictionaries(1) 
End With
```


## See also


#### Concepts


[Dictionaries Collection Object](dictionaries-object-word.md)

