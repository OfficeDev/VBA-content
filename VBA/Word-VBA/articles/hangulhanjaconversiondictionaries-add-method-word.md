---
title: HangulHanjaConversionDictionaries.Add Method (Word)
keywords: vbawd10.chm165675109
f1_keywords:
- vbawd10.chm165675109
ms.prod: word
api_name:
- Word.HangulHanjaConversionDictionaries.Add
ms.assetid: 106d6c75-5d3f-1965-79f0-942408d0450a
ms.date: 06/08/2017
---


# HangulHanjaConversionDictionaries.Add Method (Word)

Returns a  **Dictionary** object that represents a new custom spelling or conversion dictionary added to the collection of active custom spelling or conversion dictionaries.


## Syntax

 _expression_ . **Add**( **_FileName_** )

 _expression_ Required. A variable that represents a **[HangulHanjaConversionDictionaries](hangulhanjaconversiondictionaries-object-word.md)** collection.


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


[HangulHanjaConversionDictionaries Collection Object](hangulhanjaconversiondictionaries-object-word.md)

