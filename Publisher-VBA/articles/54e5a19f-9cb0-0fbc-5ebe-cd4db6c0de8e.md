
# TextRange.Script Property (Publisher)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns a  **PbFontScriptType** constant that represents the font script for a text range. Read-only.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Script**

 _expression_A variable that represents a  **TextRange** object.


### Return Value

PbFontScriptType


## Remarks
<a name="sectionSection1"> </a>

The  **Script** property value can be one of the ** [PbFontScriptType](e9bc4248-86ad-e6a8-1f50-d3ca4830118f.md)** constants declared in the Microsoft Publisher type library.


## Example
<a name="sectionSection2"> </a>

This example displays a message if the font script used in the specified text range is ASCII Latin. This example assumes that there is at least one shape on the first page of the active publication.


```
Sub DisplayScriptType() 
 If ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
 .Script = pbFontScriptAsciiLatin Then 
 MsgBox "The font script you are using is ASCII Latin." 
 End If 
End Sub
```

