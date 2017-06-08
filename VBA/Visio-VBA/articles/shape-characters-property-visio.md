---
title: Shape.Characters Property (Visio)
keywords: vis_sdr.chm11213215
f1_keywords:
- vis_sdr.chm11213215
ms.prod: visio
api_name:
- Visio.Shape.Characters
ms.assetid: dcb7fa7b-61ff-df09-8128-2d1ef4e17770
ms.date: 06/08/2017
---


# Shape.Characters Property (Visio)

Returns a  **Characters** object that represents the text of a shape. Read-only.


## Syntax

 _expression_ . **Characters**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Characters


## Remarks

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVShape.Characters**
    

## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Characters** property of a shape to get a **Characters** object. Once a **Characters** object has been retrieved, the example uses the **Shape** property of the **Characters** object to get the shape that contains the characters and demonstrates that the containing shape has been retrieved by printing its text in the Immediate window.


```vb
 
Public Sub Characters_Example() 
  
    Dim vsoOval As Visio.Shape  
    Dim vsoShapeFromCharacters As Visio.Shape  
    Dim vsoCharacters As Visio.Characters  
 
    'Create a shape and add text to it. 
    Set vsoOval = ActivePage.DrawOval(2, 5, 5, 7)  
    vsoOval.Text = "Rectangular Shape"  
 
    'Get a Characters object from the oval shape. 
    Set vsoCharacters = vsoOval.Characters  
 
    'Set the Begin and End properties so that we can 
    'replace the word "Rectangular" with "Oval" 
    vsoCharacters.Begin = 0 
    vsoCharacters.End = 11 
    vsoCharacters.Text = "Oval" 
 
    'Use the Shape property of the Characters object 
    'to get the Shape object. 
    Set vsoShapeFromCharacters = vsoCharacters.Shape  
 
    'Print the shape's text to verify that the proper Shape 
    'object was returned.  
    Debug.Print vsoShapeFromCharacters.Text 
  
End Sub
```


