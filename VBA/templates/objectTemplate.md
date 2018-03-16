---
title: Windows2 Object (Project)
keywords: vbapj.chm131358
f1_keywords:
- vbapj.chm131358
ms.prod: project-server
ms.assetid: a58383c6-12c7-81b3-10e8-81ba9180404c
ms.date: 06/08/2017
---

<!--
The example YAML block above this comment: 
title: <methodname method (workload)>
keywords: <assigned by VBA product team.>
f1_keywords: <assigned by VBA product team>
ms.prod: name of product that hosts this VBA code
ms.date: The date that the topic is checked in to master branch for publication 
-->

<!-- 
object name. For example,  Account Object (Outlook)
-->

# __objectName__ Object (__client__)

<!-- 
Description of object
For example: 
Represents a collection of windows in the application or project.

-->
Represents a ...

 


## Remarks

<!--
Describe best practices and recommendations for using the object in code.
For example:  


 **Note**  We recommend that you use the  **Windows2** collection instead of the **Windows** collection for all new development.
 

The  **Windows2** collection for the **Application** object contains all the windows in the application, whereas the **Windows2** collection for the **Project** object contains only the windows in the specified project.
 

 
Development with the .NET Framework 4, or with external components or applications that automate Project, must use the  **Windows2** object, not the **Windows** object. A Primary Interop Assembly (PIA) is required to interact correctly with the COM interface of Project if those components are to be signed. Microsoft Visual Studio generates an interop assembly based on the type library if no PIA is present, but the components then cannot be signed with a digital certificate. The PIA is installed with Project.
 -->

 

## Examples

<!-- 
Show an example of the property being read. Include description of example

For example: 
 **Using the Windows2 collection**
 

 
Use the  **[Windows2](application-windows2-property-project.md)** property to return a **Windows2** collection.
 

 
The following example cascades all the windows that are currently displayed in Project.
 

 



```
With Application.Windows2  
    For I = 1 To .Count  
        .Item(I).Activate  
        .Item(I).Top = (I - 1) * 15  
        .Item(I).Left = (I - 1) * 15  
    Next I  
End With
```

Use the  **[WindowNewWindow](application-windownewwindow-method-project.md)** method to create a new window and add it to the **Windows2** collection.
 

 
The following example creates a new window for the active project.
 

 



```
Application.WindowNewWindow
```

 **Using the Windows2 object**
 

 

## Using the Windows2 Object

Use  **Windows2** (*Index* ), where*Index* is the window index number or window caption, to return a single **Window** object.
 

 
The following example maximizes the first window in the window list.
 

 



```
Application.Windows2(1).WindowState = pjMaximized
```

The window caption is the text shown in the title bar at the top of the window when the window is not maximized. The caption is also shown in the list of open files on the bottom of the  **Windows** menu. Use the **[Caption](window-caption-property-project.md)** property to set or return the window caption. Changing the window caption does not change the name of the project.
 

 
The following example hides the window that contains the caption "Project1".
 

 



```
If Application.Windows2(1).Caption = "Project1" Then  
    Application.Windows2(1).Visible = False  
End If
```


-->


## Properties

<!-- 
Link to the properties exposed by the object

For example: 



|**Name**|
|:-----|
|[ActiveWindow](windows2-activewindow-property-project.md)|
|[Application](windows2-application-property-project.md)|
|[Count](windows2-count-property-project.md)|
|[Item](windows2-item-property-project.md)|
|[Parent](windows2-parent-property-project.md)|

-->
## Methods

<!-- 
Link to the methods exposed by the object

For example: 



|**Name**|
|:-----|
|[refresh](window-refresh-method-project.md)|


-->


## See also
<!-- 
Optional:  Link to relevant API or conceptual articles
-->


#### Other resources

<!-- 
Link to VBA client object model article
For example: 

 
[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
-->
