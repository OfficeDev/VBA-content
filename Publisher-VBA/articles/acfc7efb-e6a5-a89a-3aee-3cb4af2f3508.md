
# Application Object (Publisher)

 **Last modified:** July 28, 2015

Represents the Microsoft Publisher application. The  **Application** object includes properties and methods that return top-level objects. For example, the **ActiveDocument** property returns a **Document** object.

## Remarks

When using Microsoft Visual Basic for Applications in Publisher, all of the properties and methods of the  **Application** object can be used without the **Application** object qualifier. For example, instead of typing `Application.ActiveDocument.PrintOut`, you can type  `ActiveDocument.PrintOut`. Properties and methods that can be used without the  **Application** object qualifier are considered "global." To view the global properties and methods in the Object Browser, click **&lt;globals&gt;** at the top of the list in the **Classes** box. When accessing the Publisher object model from a non-Publisher project, all properties and methods must be fully qualified.


## Example

Use the  ** [Application](f3ed5997-b8ef-4729-4537-ae21424d2007.md)**property to return the  **Application** object. The following example displays the application name.


```
Sub ShowAppName() 
 MsgBox Application.Name 
End Sub
```


## See also


#### Other resources


 [Application Object Members](aa4d515b-f779-b8b5-968a-8e5f7466fb56.md)
