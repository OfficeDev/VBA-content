
# Style.BuiltIn Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


 **True** if the specified object is one of the built-in styles or caption labels in Word. Read-only **Boolean**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **BuiltIn**

 _expression_A variable that represents a  ** [Style](473f8f41-2cba-769e-c0da-441d9d85b009.md)** object.


## Remarks
<a name="sectionSection1"> </a>

You can specify built-in styles across all languages by using the  **WdBuiltinStyle** constants or within a language by using the style name for the language version of Word. For example, if you specify U.S. English in your Microsoft Office language settings, the following statements are equivalent:


```
ActiveDocument.Styles(wdStyleHeading1) 
ActiveDocument.Styles("Heading 1")
```


## Example
<a name="sectionSection2"> </a>

This example checks all the styles in the active document. When it finds a style that isn't built in, it displays the name of the style.


```
Dim styleLoop As Style 
 
For Each styleLoop in ActiveDocument.Styles 
 If styleLoop.BuiltIn = False Then 
 Msgbox styleLoop.NameLocal 
 End If 
Next styleLoop
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Style Object](473f8f41-2cba-769e-c0da-441d9d85b009.md)
#### Other resources


 [Style Object Members](37c68e72-c745-bc9c-1547-0cf177cbdef4.md)
