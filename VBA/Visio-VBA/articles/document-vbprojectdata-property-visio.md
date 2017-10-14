---
title: Document.VBProjectData Property (Visio)
keywords: vis_sdr.chm10550925
f1_keywords:
- vis_sdr.chm10550925
ms.prod: visio
api_name:
- Visio.Document.VBProjectData
ms.assetid: dca456ea-dc82-0092-35d1-68b95d51e0b2
ms.date: 06/08/2017
---


# Document.VBProjectData Property (Visio)

Returns the Microsoft Visual Basic project data stored with a document. Read-only.


## Syntax

 _expression_ . **VBProjectData**

 _expression_ A variable that represents a **Document** object.


### Return Value

Byte()


## Example



You can use the  **VBProjectData** property to determine whether a document has a project. The following macro shows how to get a reference to a document in Microsoft Visio to determine whether the document has a project. The code runs from a program outside the Visio document.




```vb
Private Sub Form_Load() 
 
 'Declare document variable 
 'and Array variable to hold project data. 
 Dim vsoDocument As Visio.Document 
 Dim btProjectData() As Byte 
 
 'Get the first object in the Documents collection 
 'of this instance of Visio. 
 Set vsoDocument = GetObject(, "Visio.Application").Documents(1) 
 
 'Populate the array with project data. 
 btProjectData = vsoDocument.VBProjectData 
 Debug.Print LBound(btProjectData); UBound(btProjectData) 
 
End Sub
```

If the document had no project associated with it, "0 -1" would be reported in the Immediate window. If the document had a project, the upper bound would be some number greater than zero (0). For example, "0 1535" would indicate that a project had 1,536 bytes of data.


