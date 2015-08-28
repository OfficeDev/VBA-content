
# Presentation.AddToFavorites Method (PowerPoint)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Adds a shortcut that represents the current selection in the specified presentation to the Windows Favorites folder.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **AddToFavorites**

 _expression_A variable that represents a  **Presentation** object.


## Remarks
<a name="sectionSection1"> </a>

The shortcut name is the display name of the document, if that's available; otherwise, the shortcut name is as calculated in HLINK.DLL.


## Example
<a name="sectionSection2"> </a>

This example adds a hyperlink to the active presentation to the Favorites folder in the Windows program folder.


```
Application.ActivePresentation.AddToFavorites
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Presentation Object](ec75cf52-69f8-d35b-0a26-4a8da8a9683f.md)
#### Other resources


 [Presentation Object Members](b3538c7e-5fd9-d34d-ab5c-0105dbd516d0.md)
