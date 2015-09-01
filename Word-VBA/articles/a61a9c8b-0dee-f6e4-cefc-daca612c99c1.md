
# Document.CheckSpelling Method (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Begins a spelling check for the specified document or range. .


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **CheckSpelling**( **_CustomDictionary_**,  **_IgnoreUppercase_**,  **_AlwaysSuggest_**,  **_CustomDictionary2_**,  **_CustomDictionary3_**,  **_CustomDictionary4_**,  **_CustomDictionary5_**,  **_CustomDictionary6_**,  **_CustomDictionary7_**,  **_CustomDictionary8_**,  **_CustomDictionary9_**,  **_CustomDictionary10_**)

 _expression_Required. A variable that represents a  ** [Document](8d83487a-2345-a036-a916-971c9db5b7fb.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|IgnoreUppercase|Optional| **Variant**| **True** if capitalization is ignored. If this argument is omitted, the current value of the **IgnoreUppercase** property is used.|
|AlwaysSuggest|Optional| **Variant**| **True** for Microsoft Word to always suggest alternative spellings. If this argument is omitted, the current value of the **SuggestSpellingCorrections** property is used.|

## Remarks
<a name="sectionSection1"> </a>

If the document or range contains errors, this method displays the  **Spelling and Grammar** dialog box ( **Tools** menu), with the **Check grammar** check box cleared. For a document, this method checks all available stories (such as headers, footers, and text boxes).


## Example
<a name="sectionSection2"> </a>

The following example checks the spelling in the active document.


```
ActiveDocument.CheckSpelling
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Document Object](8d83487a-2345-a036-a916-971c9db5b7fb.md)
#### Other resources


 [Document Object Members](fc9ab457-0888-f917-3d52-387168ac23b9.md)
