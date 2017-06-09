---
title: Project.MakeServerURLTrusted Method (Project)
keywords: vbapj.chm132401
f1_keywords:
- vbapj.chm132401
ms.prod: project-server
api_name:
- Project.Project.MakeServerURLTrusted
ms.assetid: 8ef5ae1c-f22f-325c-07a9-253e64c62cb0
ms.date: 06/08/2017
---


# Project.MakeServerURLTrusted Method (Project)

Adds the URL specified in the  **[ServerURL](http://msdn.microsoft.com/library/a204c795-73a3-4ce2-a582-3afd951914c7%28Office.15%29.aspx)** property to the **Trusted sites** zone in the **Security** tab of the **Internet Options** dialog box in Internet Explorer.


## Syntax

 _expression_. **MakeServerURLTrusted**

 _expression_ A variable that represents a **Project** object.


## Remarks

If no Project Server URL is specified for the project, Project Professional displays an error message stating: "A Project Server URL has not been specified. To specify a URL, on the Tools menu, click Options, and then click the Collaboration tab."


## Example

The following sample adds the URL specified in  **Collaboration Options** ( **Collaborate** menu) to the list of trusted sites in Internet Explorer. Upon confirmation, Project switches to a **Resource Sheet** view and displays the displays the **Build Team for <Project Name>** dialog box when connected to Project Server .


```vb
Sub MakeURLTrusted() 
   If Projects.Count = 0 Then 
      MsgBox "You must have at least one active project open." 
      Exit Sub 
   End If 
 
   If ActiveProject.ServerURL = "" Then 
      MsgBox "A Project Server URL has not been " _ 
         &; "specified." &; Chr(13) &; "Click OK, and then " _
         &; "specify a valid URL in the Project Server Accounts dialog box." 
   Else 
      ActiveProject.MakeServerURLTrusted 
      ViewApply Name:="Resource Sheet" 
      Application.AddResourcesFromProjectServer 
   End If 
End Sub
```


