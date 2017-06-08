---
title: Document.AlternateNames Property (Visio)
keywords: vis_sdr.chm10513085
f1_keywords:
- vis_sdr.chm10513085
ms.prod: visio
api_name:
- Visio.Document.AlternateNames
ms.assetid: 2d0a3f45-e9b4-385b-23c9-2a0a70375202
ms.date: 06/08/2017
---


# Document.AlternateNames Property (Visio)

Gets or sets the alternate names for a document. Read/write.


## Syntax

 _expression_ . **AlternateNames**

 _expression_ A variable that represents a **Document** object.


### Return Value

String


## Remarks

The application stores document names in the following situations:


- Templates store stencil names. For example, the  **Basic Flowchart** template stores the names of the **Basic Flowchart Shapes.vss** and **Backgrounds.vss** stencils. These stencils are opened with the **Basic Flowchart** template.
    
- Master shortcuts store stencil names. For example, a shortcut for the  **Data** shape stores the name of the stencil on which the **Data** shape is stored— **Basic Flowchart Shapes.vss**.
    
When the application opens a document or accesses the  **Document** object's collection, it uses the document name. If Microsoft Visio can't find the document name, it looks for alternate names for those stencils that are in the correct path. (To add a path, click the **File** tab, click **Options**, click  **Advanced**, and then, under  **General**, click ** File Locations**.) For example, suppose you created the stencil named "New Shapes 2008.vss." The following year you revised the stencil and renamed it "New Shapes 2009.vss." Any templates that opened  **New Shapes 2008.vss** should now open **New Shapes 2009.vss**. To do this, set the  **AlternateNames** property of **New Shapes 2009.vss** to "New Shapes 2008.vss." The following Microsoft Visual Basic code shows one way to do this:




```
Visio.Documents("New Shapes 2009.vss").AlternateNames = "New Shapes 2008.vss"
```

The  **AlternateNames** property is empty until you set it by using Automation. Each of the alternate names in the string should contain the file name, but no folder information. You can also include comments in angle brackets (<>), because the application ignores anything in angle brackets. For example, you could use the following code to set the **AlternateNames** property:




```
Visio.Documents("HRShapes.vss").AlternateNames = "Human Resources Shapes.vss; <old name> HRDept Shapes.vss"
```


## Example

The following macro shows how to get and set the  **AlternateNames** property of the current document. It demonstrates that the property is empty until you set it.


```vb
 
Public Sub AlternateNames_Example() 
  
    'Get the AlternateNames property of the document.  
    Debug.Print "Alternate name is: "; ThisDocument.AlternateNames 
 
    'Set the AlternateNames property of the document.  
    ThisDocument.AlternateNames = "Test Shapes.vss"  
    Debug.Print "Alternate name is: "; ThisDocument.AlternateNames  
 
End Sub
```


