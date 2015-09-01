
# Selection.DetectLanguage Method (Word)

 **Last modified:** July 28, 2015

Analyzes the specified text to determine the language that it is written in.

## Syntax

 _expression_. **DetectLanguage**

 _expression_Required. A variable that represents a  ** [Selection](7b574a91-c33e-ecfd-6783-6b7528b2ed8f.md)** object.


## Remarks

The results of the  **DetectLanguage** method are stored in the **LanguageID**property on a character-by-character basis. To read the  ** [LanguageID](8af15ba5-19f0-2a65-e44a-a9fed55f8239.md)** property, you must first specify a selection or range of text.

If a selection contains a partial sentence, the selection is extended to the end of the sentence.

If the  **DetectLanguage** method has already been applied to the specified text, the **LanguageDetected**property is set to  **True**. To reevaulate the language of the specified text, you must first set the  ** [LanguageDetected](18eba980-a599-e6f0-7d73-bee6da0474be.md)** property to **False**.


## See also


#### Concepts


 [Selection Object](7b574a91-c33e-ecfd-6783-6b7528b2ed8f.md)
#### Other resources


 [Selection Object Members](71e67a43-d40a-ad9a-8ef2-c5c487733e0d.md)
