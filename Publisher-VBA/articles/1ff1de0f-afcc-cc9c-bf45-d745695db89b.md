
# ParagraphFormat.ListBulletFontSize Property (Publisher)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Sets or retrieves a  **Single** that represents the list bullet font size from the specified paragraphs. Read/write.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **ListBulletFontSize**

 _expression_A variable that represents a  **ParagraphFormat** object.


### Return Value

Single


## Remarks
<a name="sectionSection1"> </a>

Returns an "Access Denied" message if the list is not a bulleted list.


## Example
<a name="sectionSection2"> </a>

This example tests to see if the list type is a bulleted list. If it is, the  **ListFontSize** is set to 24 and the **ListBulletFontName** is set to "Verdana".


```
Dim objParaForm As ParagraphFormat 
 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
.TextFrame.TextRange.ParagraphFormat 
 
With objParaForm 
 If .ListType = pbListTypeBullet Then 
 .ListBulletFontSize = 24 
 .ListBulletFontName = "Verdana" 
 End If 
End With 
 
 

```

