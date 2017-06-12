---
title: Application.AddResourcesFromProjectServer Method (Project)
keywords: vbapj.chm2130
f1_keywords:
- vbapj.chm2130
ms.prod: project-server
api_name:
- Project.Application.AddResourcesFromProjectServer
ms.assetid: 74fe4224-0019-5daa-11ae-3bdd6f2f5abb
ms.date: 06/08/2017
---


# Application.AddResourcesFromProjectServer Method (Project)

Opens the ** Build Team** dialog box if you are connected to Project Server and are currently in a resource view.


## Syntax

 _expression_. **AddResourcesFromProjectServer**

 _expression_ A variable that represents an **Application** object.


### Return Value

 **Boolean**


## Remarks

The  **AddResourcesFromProjectServer** method is only available in resource views and returns a trappable error (error code 1100) when applied in a non-resource view.


## Example

The following example displays the  **Build Team from Project Server** dialog box. First, Project adds the URL specified in **Collaboration Options** ( **Collaborate** menu) to Microsoft Internet Explorer's list of trusted sites. Upon confirmation, Project switches to a **Resource Sheet** view and displays the ** Build Team from Project Server** dialog box, if connected to My Computer in workgroup mode. Project displays the ** Build Team from <Project Name>** dialog box when connected to Project Server.


```vb
Sub AddResources() 
   If Projects.Count = 0 Then 
      MsgBox "You must have at least one active project open." 
      Exit Sub 
   End If 
 
   If ActiveProject.ServerURL = "" Then 
      MsgBox "A Project Server URL has not been " _ 
         &; "specified." &; Chr(13) &; "Click OK to select " _ 
         &; "'Collaborate Using Project Server' and " _ 
         &; "specify a valid URL in the Options dialog box " _ 
         &; "(Tools menu)." 
      Application.OptionsWorkgroup 
   Else 
      ActiveProject.MakeServerURLTrusted 
      ViewApply Name:="Resource Sheet" 
      Application.AddResourcesFromProjectServer 
   End If 
End Sub
```


