---
title: Project.CustomDocumentProperties Property (Project)
ms.prod: project-server
api_name:
- Project.Project.CustomDocumentProperties
ms.assetid: 49e532bc-4bc2-c9e7-c6d0-253540572093
ms.date: 06/08/2017
---


# Project.CustomDocumentProperties Property (Project)

Gets a  **DocumentProperties** collection representing the custom properties of the document. Read-only **Object**.


## Syntax

 _expression_. **CustomDocumentProperties**

 _expression_ A variable that represents a **Project** object.


## Remarks

For more information, see  _DocumentProperties Collection Object_ in the Microsoft Office Visual Basic Reference.

To use this property, you must include a reference to the Microsoft Office 14.0 Object Library by using the  **References** command on the **Tools** menu. The Object Library contains definitions for the Visual Basic objects, properties, methods, and constants used to manipulate document properties.

Use the  **BuiltinDocumentProperties** property to return the collection of built-in document properties.


## Example

In the following example, the  **Date completed** custom property value would be **Nothing** if the property is added to the project, but the project is not completed. Before you run the **TestDocProps** example, add some tasks to the active project and assign them to a resource.


```vb
Sub TestDocProps()
    Dim docProps As Office.DocumentProperties
    Dim docProp As Office.DocumentProperty
    Dim numProps As Integer
    
    Set docProps = ActiveProject.CustomDocumentProperties
    
    numProps = docProps.Count
    Debug.Print "Number of custom document properties: " &; numProps
    
    For Each docProp In docProps
        If (docProp.Name = "Date completed") Then
            Debug.Print "Date completed: (none) "
        Else
            Debug.Print docProp.Name &; vbTab &; ": " &; docProp.Value
        End If
    Next docProp
End Sub
```

Following are the results of the  **TestDocProps** macro, for a project that does not have the **Date completed** property added:




```
Number of custom document properties: 7
% Complete  : 0%
Cost    : $0.00
Duration    : 5 days?
Finish  : Thu 5/7/09
Start   : Fri 5/1/09
Work    : 40h
% Work Complete : 0%
```


