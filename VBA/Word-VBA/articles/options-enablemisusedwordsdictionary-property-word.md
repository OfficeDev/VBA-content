---
title: Options.EnableMisusedWordsDictionary Property (Word)
keywords: vbawd10.chm162988370
f1_keywords:
- vbawd10.chm162988370
ms.prod: word
api_name:
- Word.Options.EnableMisusedWordsDictionary
ms.assetid: 53ec74bd-d1fb-39ee-6ddb-4cca840c13dd
ms.date: 06/08/2017
---


# Options.EnableMisusedWordsDictionary Property (Word)

 **True** if Microsoft Word checks for misused words when checking the spelling and grammar in a document. Read/write **Boolean** .


## Syntax

 _expression_ . **EnableMisusedWordsDictionary**

 _expression_ A variable that represents a **[Options](options-object-word.md)** object.


## Remarks

Word looks for the following when checking for misused words: incorrect usage of adjectives and adverbs, comparatives and superlatives, "like" as a conjunction, "nor" versus "or," "what" versus "which," "who" versus "whom," units of measurement, conjunctions, prepositions, and pronouns.


## Example

This example sets Word to ignore misused words when checking spelling and grammar.


```vb
Options.EnableMisusedWordsDictionary = False
```


## See also


#### Concepts


[Options Object](options-object-word.md)

