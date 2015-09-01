
# Options.EnableMisusedWordsDictionary Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


 **True** if Microsoft Word checks for misused words when checking the spelling and grammar in a document. Read/write **Boolean**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **EnableMisusedWordsDictionary**

 _expression_A variable that represents a  ** [Options](873b7b99-3fe1-fd89-9ece-a9355cb827dc.md)** object.


## Remarks
<a name="sectionSection1"> </a>

Word looks for the following when checking for misused words: incorrect usage of adjectives and adverbs, comparatives and superlatives, "like" as a conjunction, "nor" versus "or," "what" versus "which," "who" versus "whom," units of measurement, conjunctions, prepositions, and pronouns.


## Example
<a name="sectionSection2"> </a>

This example sets Word to ignore misused words when checking spelling and grammar.


```
Options.EnableMisusedWordsDictionary = False
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Options Object](873b7b99-3fe1-fd89-9ece-a9355cb827dc.md)
#### Other resources


 [Options Object Members](76cd9dfe-6bbb-4c3d-0bfc-79a62bedd15e.md)
