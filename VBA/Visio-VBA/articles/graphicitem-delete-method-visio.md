---
title: GraphicItem.Delete Method (Visio)
keywords: vis_sdr.chm16951425
f1_keywords:
- vis_sdr.chm16951425
ms.prod: visio
api_name:
- Visio.GraphicItem.Delete
ms.assetid: cd395089-594a-a021-0455-5bd7de9c3468
ms.date: 06/08/2017
---


# GraphicItem.Delete Method (Visio)

Deletes a  **GraphicItem** object from the **GraphicItems** collection of a **Master** object of type **visTypeDataGraphic** .


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **Delete**

 _expression_ A variable that represents a **GraphicItem** object.


### Return Value

Nothing


## Remarks

Before you can delete a graphic item, you must use the  **[Master.Open](master-open-method-visio.md)** method to open for editing a copy of the data graphic master whose **GraphicItems** collection the graphic item belongs to. After you have deleted the graphic item and made whatever other edits you want to make, use the **Master.Close** method to commit changes.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Delete** method to delete an existing graphic item from the **GraphicItems** collection of a **Master** object. It deletes the graphic item most recently added to the collection and prints the count of graphic items in the collection of the master copy, both before and after the deletion, to the **Immediate** window. After it closes the master copy, it also prints the count of graphic items in the collection of the master itself, to show that actions performed on the copy get committed to the master.

The macro assumes that there is an existing data graphic master in your project in whose  **GraphicItems** collection has at least one member. You can determine the name of an existing data graphic master by moving your mouse over the master in the **Data Graphics** task pane in the Visio user interface. The master in this example is named "Data Graphic."




```vb
Public Sub Delete_Example() 
 
    Dim vsoMaster As Visio.Master 
    Dim vsoMasterCopy As Visio.Master 
    Dim intGraphicItemCount As Integer 
 
    Set vsoMaster = ActiveDocument.Masters("Data Graphic") 
    Set vsoMasterCopy = vsoMaster.Open 
     
    intGraphicItemCount = vsoMasterCopy.GraphicItems.Count 
     
    Debug.Print "Before delete", intGraphicItemCount 
    vsoMasterCopy.GraphicItems(intGraphicItemCount).Delete 
    Debug.Print "After delete", vsoMasterCopy.GraphicItems.Count 
    vsoMasterCopy.Close 
    Debug.Print "After close", vsoMaster.GraphicItems.Count 
     
End Sub
```


