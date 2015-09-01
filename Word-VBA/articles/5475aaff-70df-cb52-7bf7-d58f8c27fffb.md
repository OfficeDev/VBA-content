
# DefaultWebOptions.CheckIfOfficeIsHTMLEditor Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


 **True** if Microsoft Word checks to see whether an Office application is the default HTML editor when you start Word. Read/write **Boolean**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **CheckIfOfficeIsHTMLEditor**

 _expression_A variable that represents a  ** [DefaultWebOptions](7459af1e-c495-f84f-929c-f7a611ec49b3.md)** object.


## Remarks
<a name="sectionSection1"> </a>

The  **CheckIfOfficeIsHTMLEditor** property returns **False** if Word does not perform this check. The default value is **True**.

 This property is used only if the Web browser you are using supports HTML editing and HTML editors. To use a different HTML editor, you must set this property to **False** and then register the editor as the default system HTML editor.


## Example
<a name="sectionSection2"> </a>

This example causes Microsoft Word not to check to see whether an Office application is the default HTML editor.


```
Application.DefaultWebOptions _ 
 .CheckIfOfficeIsHTMLEditor = False
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [DefaultWebOptions Object](7459af1e-c495-f84f-929c-f7a611ec49b3.md)
#### Other resources


 [DefaultWebOptions Object Members](2ec195b5-f843-6a29-9070-a86a7ff1d7fc.md)
