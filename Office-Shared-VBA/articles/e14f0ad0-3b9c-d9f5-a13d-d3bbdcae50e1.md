
# TextRange2.Replace Method (Office)

 **Last modified:** July 28, 2015

Finds specific text in a text range, replaces the found text with a specified string, and returns a  **TextRange2** object that represents the first occurrence of the found text. Returns **Nothing** if no match is found.

## Syntax

 _expression_. **Replace**( **_FindWhat_**,  **_ReplaceWhat_**,  **_After_**,  **_MatchCase_**,  **_WholeWords_**)

 _expression_An expression that returns a  **TextRange2** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|FindWhat|Required| **String**|The text to search for.|
|ReplaceWhat|Required| **String**|The text you want to replace the found text with.|
|After|Optional| **Long**|The position of the character (in the specified text range) after which you want to search for the next occurrence of  **FindWhat**. For example, if you want to search from the fifth character of the text range, specify 4 for  **After**. If this argument is omitted, the first character of the text range is used as the starting point for the search.|
|MatchCase|Optional| **MsoTriState**|Determines whether a distinction is made on the basis of case.|
|WholeWords|Optional| **MsoTriState**|Determines whether only whole words are searched.|

### Return Value

TextRange2


## See also


#### Concepts


 [TextRange2 Object](a6a59c9b-9b64-c1e2-2e98-a1f99025c877.md)
#### Other resources


 [TextRange2 Object Members](26daffff-b9ef-fd94-f5b7-ed3a09840cb6.md)
