
# ParagraphFormat.TextDirection Property (Publisher)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets a  **PbTextDirection** constant indicating the direction in which text flows in the specified paragraph. Read/write.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **TextDirection**

 _expression_A variable that represents a  **ParagraphFormat** object.


### Return Value

PbTextDirection


## Remarks
<a name="sectionSection1"> </a>

This property is meant to be used in conjunction with documents that have text in both left-to-right and right-to-left languages. Setting the property to a value that is not in accordance with the text direction dictated by the language in use may have unpredictable results.

The  **TextDirection** property value can be one of the **PbTextDirection** constants declared in the Microsoft Publisher type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **pbTextDirectionLeftToRight**| Text flows from left to right.|
| **pbTextDirectionMixed**|Return value indicating a range containing some left-to-right text and some right-to-left text.|
| **pbTextDirectionRightToLeft**|Text flows from right to left.|

## Example
<a name="sectionSection2"> </a>

The following example changes the text direction of the first shape on page one so that it flows from right-to-left.


```
ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
 .ParagraphFormat.TextDirection = pbTextDirectionRightToLeft
```

