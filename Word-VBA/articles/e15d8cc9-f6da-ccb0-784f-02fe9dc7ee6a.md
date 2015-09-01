
# Application.DefaultSaveFormat Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets the default format that will appear in the  **Save as type** box in the **Save As** dialog box. Read/write **String**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **DefaultSaveFormat**

 _expression_An expression that represents a  ** [Application](d1cf6f8f-4e88-bf01-93b4-90a83f79cb44.md)** object.


## Remarks
<a name="sectionSection1"> </a>

The string used with this property is the file converter class name. The class names for internal Word formats are listed in the following table.



|**Word format**|**File converter class name**|
|:-----|:-----|
|Word Document|""|
|Document Template|"Dot"|
|Text Only|"Text"|
|Text Only with Line Breaks|"CRText"|
|MS-DOS Text|"8Text"|
|MS-DOS Text with Line Breaks|"8CRText"|
|Rich Text Format|"Rtf"|
|Unicode Text|"Unicode"|
Use the  ** [ClassName](71124adf-11fc-e42d-a9f5-940f7fea97af.md)**property of the  ** [FileConverter](41af2a9b-75cc-253d-4954-4fb42c88530f.md)** object to determine the class name of an external file converter.


## Example
<a name="sectionSection2"> </a>

This example sets the Word document format as the default save format.


```
Application.DefaultSaveFormat = ""
```

This example returns the current setting that Word uses for saving a document.




```
Msgbox Application.DefaultSaveFormat
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Application Object](d1cf6f8f-4e88-bf01-93b4-90a83f79cb44.md)
#### Other resources


 [Application Object Members](71669f1e-65f1-b0f1-b67d-355dfdbebe50.md)
