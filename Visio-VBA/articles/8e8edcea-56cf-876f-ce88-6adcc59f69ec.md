
# VisWebPageSettings.TargetPath Property (Visio Save As Web)

 **Last modified:** June 07, 2012

 _**Applies to:** Visio 2013_

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Specifies the path where the Web page and its supporting files are placed. Read/write.

## Syntax
<a name="sectionSection0"> </a>

 _expression_. **TargetPath**

 _expression_An expression that returns a  ** [VisWebPageSettings](http://msdn.microsoft.com/library/14280ea7-e8b1-d4b2-941b-121f2c17f787%28Office.15%29.aspx)** object.


### Return Value

 **String**


## Remarks
<a name="sectionSection1"> </a>

When you save a drawing as a Web page, you must use the  **TargetPath** property to supply the full target path. The **TargetPath** property is reset to a null value after each export: each time you save a drawing as a Web page you must explicitly supply the target path. In addition, the **TargetPath** value is not persisted between sessions of Visio.

The value of the **TargetPath** property corresponds to the folder name and file name selected in the **Save As** dialog box (click the **BackstageButton** tab, and then click **Save As**).


## Example
<a name="sectionSection2"> </a>

The following macro shows how to save the active document as a Web page and place the resulting HTML file and supporting files as flat files in the  _targetpath_ folder. Because the ** [StoreInFolder](ed0cf76a-a68d-cfa7-538c-91df5234a0d0.md)** property is set to **False**, the supporting files are placed in the same folder as the root HTML file, instead of in a separate folder that has the name  _filename_files_ or _filename.files_, depending on the language.


```
Public Sub TargetPath_Example()
    Dim vsoSaveAsWeb As VisSaveAsWeb 
    Dim vsowebSettings As VisWebPageSettings

    Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
    Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings
 
    With vsoWebSettings
         .StoreInFolder = False
         .TargetPath = "<variable>targetpath\filename.htm</variable>"
    End With
 
    vsoSaveAsWeb.CreatePages 
End Sub

```

