---
title: Document.VBProject Property (Visio)
keywords: vis_sdr.chm10514635
f1_keywords:
- vis_sdr.chm10514635
ms.prod: visio
api_name:
- Visio.Document.VBProject
ms.assetid: 087e9cdc-c21d-6f02-05ce-4c3fa6e09cff
ms.date: 06/08/2017
---


# Document.VBProject Property (Visio)

Returns an automation object that you can use to control the Microsoft Visual Basic for Applications (VBA) project of the document. Read-only.


## Syntax

 _expression_ . **VBProject**

 _expression_ A variable that represents a **Document** object.


### Return Value

Object


## Remarks

To get information about the object returned by the  **VBProject** property, follow these steps:


### To get information about the object returned by the VBProject property


1. In the  **Code** group on the[Developer](http://msdn.microsoft.com/library/1bdc55f5-8fc7-7257-03d5-c049eceb29ff%28Office.15%29.aspx) tab, click **Visual Basic**.
    
2. In the Visual Basic Editor, on the  **Tools** menu, click **References**.
    
3. In the  **References** dialog box, click **Microsoft Visual Basic for Applications Extensibility 5.3**, and then click  **OK**.
    
4. On the  **View** menu, click **Object Browser**.
    
5. In the  **Project/Library** list, select the **VBIDE** type library.
    
6. In the  **Classes** list, examine the class named **VBProject**.
    
If no VBA project already exists in the document, the  **VBProject** property creates one.

Beginning with Visio 2002, the  **VBProject** property raises an exception if you are running in a security-enhanced environment and your system administrator has blocked access to the Visual Basic object model. There is no user interface or programmatic way to turn this on—the system administrator must turn on (or off) access by setting a Group Policy. This helps protect against viruses that spread by accessing the Visual Basic projects in commonly used templates and injecting the virus code into them.


## Example

This VBA macro shows how to print the names of libraries referenced by a VBA project in the Immediate window.

Before running this code, make sure the  **Trust access to the VBA project object model** check box is selected under **Developer Macro Settings** on the **Macro Settings** page of the **Trust Center** dialog box (click the **File** tab, click **Options**, click  **Trust Center**, and then click  **Trust Center Settings**). 




```vb
Public Sub VBProject_Example()  
 
    Dim varThisProject As Variant 
    Dim intReferences As Integer 
 
    Set varThisProject = ThisDocument.VBProject  
 
    intReferences = varThisProject.References.Count  
    While intReferences > 0  
        Debug.Print varThisProject.References(intReferences).Name  
        intReferences = intReferences - 1  
    Wend 
 
End Sub
```


