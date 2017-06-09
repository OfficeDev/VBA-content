---
title: Master.Open Method (Visio)
keywords: vis_sdr.chm10751195
f1_keywords:
- vis_sdr.chm10751195
ms.prod: visio
api_name:
- Visio.Master.Open
ms.assetid: 3f14f3b2-1cfb-ccf9-b344-7fbf80ae9a26
ms.date: 06/08/2017
---


# Master.Open Method (Visio)

Opens an existing master so that it can be edited.


## Syntax

 _expression_ . **Open**

 _expression_ A variable that represents a **Master** object.


### Return Value

Master


## Remarks

You can use the  **Open** method for a **Master** object in conjunction with the **Close** method to reliably edit the shapes and cells of a master. In some previous versions of Visio, you could edit a **Master** object's shapes and cells, but the changes were not pushed to instances of the master, and alignment box information displayed when instancing the edited master was not correct.


### To edit the shapes and cells of a Master object from a program




1. Open the  **Master** object for editing by using _masterObjCopy_ = _masterObj_ . **Open** . This code fails if there is a drawing window open into _masterObj_ or if other programs already have _masterObj_ open. If the **Open** method succeeds, _masterObjCopy_ is a copy of _masterObj_ .
    
2. Change any shapes and cells in  _masterObjCopy_ , not _masterObj_ .
    
3. Close the  **Master** object by using _masterObjCopy_ . **Close** . The **Close** method fails if _masterObjCopy_ isn't a **Master** object that resulted from a prior _masterObj_ . **Open** call. Otherwise, the **Close** method merges the changes made in step 2 from _masterObjCopy_ back into _masterObj_ . It also updates all instances of _masterObj_ to reflect the changes and update information cached in _masterObj_ . If _masterObj_ . **IconUpdate** isn't **visManual** (0), the **Close** method updates the icon shown in the stencil window for _masterObj_ to depict an image of _masterObjCopy_ .
    
If you change the shapes and cells of a master directly, as opposed to opening and closing the master as described in the procedure above, the effects listed in step 3 don't occur.

A program that creates a copy of a  _masterObj_ for editing should both close and release the copy. Microsoft Visual Basic typically releases it automatically. However, when you are coding in C or C++, you must explicitly release the copy, just as you would for any other object.


 **Note**  Starting with Microsoft Office Visio 2003, only user-created stencils are editable. By default, Visio stencils are not editable. 


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to open a  **Master** object for editing. It opens a copy of a master from the document stencil and changes the fill foreground color of the master and all shapes in the drawing derived from the master.

Before running this macro, close all open Visio documents. Then, click the  **File** tab, click **New**, and then click  **Create** to open a new document based on no template. Click the **Rectangle** tool, and draw a rectangle on the drawing page. Open the document stencil (in the **Shapes** window, click **More Shapes**, click  **Show Document Stencil**), and then drag the rectangle shape onto the document stencil to create a master. Finally, drag several copies of the rectangle master onto the drawing page.




```vb
 
Public Sub OpenMaster_Example() 
 
    Dim vsoMaster As Visio.Master 
    Dim vsoMasterCopy As Visio.Master 
    Dim vsoShape As Visio.Shape 
    Dim vsoCell As Visio.Cell 
 
    Set vsoMaster = Visio.Documents.Masters(1) 
    Set vsoMasterCopy = vsoMaster.Open 
 
    Set vsoShape = vsoMasterCopy.Shapes.Item(1) 
 
    Set vsoCell = vsoShape.CellsU("FillForegnd") 
    vsoCell.Formula = 9 
 
    vsoMasterCopy.Close 
 
End Sub
```


