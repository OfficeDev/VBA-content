---
title: GraphicItems.AddCopy Method (Visio)
keywords: vis_sdr.chm16860420
f1_keywords:
- vis_sdr.chm16860420
ms.prod: visio
api_name:
- Visio.GraphicItems.AddCopy
ms.assetid: 9956c5a5-8200-4e2a-c219-0a26fc40b414
ms.date: 06/08/2017
---


# GraphicItems.AddCopy Method (Visio)

Adds a copy of a  **GraphicItem** object to the **GraphicItems** collection of a **Master** object of type **visTypeDataGraphic** .


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **AddCopy**( **_GraphicItem_** )

 _expression_ An expression that returns a **GraphicItems** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _GraphicItem_|Required| **[IVGRAPHICITEM]**|The  **GraphicItem** object to copy.|

### Return Value

GraphicItem


## Remarks

The  **GraphicItem** object you want to add must already exist in the **GraphicItems** collection of another **Master** object of type **visTypeDataGraphic** .

After you use the  **Master.Open** to open a copy of a **Master** object of type **visTypeDataGraphic** for editing, you must use the **[Master.Close](master-close-method-visio.md)** method to commit any changes you made to the master while it was open. Closing a copy of a data-graphic master also reapplies the data graphic to all shapes to which it was previously applied.


 **Note**  For more information about why it is necessary to edit a copy of a master instead of the master itself, see  **[Master.Open](master-open-method-visio.md)** .


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **AddCopy** method to add a copy of an existing data-graphic item to the **GraphicItems** collection of a **Master** object.

The macro uses the  **Masters.AddEx** method to add a new **Master** object of type **visTypeDataGraphic** to the **Masters** collection of the active document. Then, it uses the **Master.Open** method to get a copy of the new data graphic master to edit.

Next, the method gets a copy of an existing data-graphic item that is the first item in the  **GraphicItems** collection of an existing master. Finally, it uses the **AddCopy** method to add the copy of the existing data-graphic item to the **GraphicItems** collection of the copy of the new master, and then closes the master copy.

The macro assumes that you know the name of the existing data-graphic master (" _old_master_name_ ") that contains one or more graphic items you want to add to the new master. You can determine the name of an existing data graphic master by moving your mouse over the master in the **Data Graphics** task pane in the Visio user interface.




```vb
Public Sub AddCopy_Example() 
 
    Dim vsoMaster As Visio.Master 
    Dim vsoMasterCopy As Visio.Master 
    Dim vsoMaster_Old As Visio.Master 
    Dim vsoGraphicItem As GraphicItem 
    Dim vsoGraphicItem_Old As Visio.GraphicItem 
 
    Set vsoMaster = Visio.ActiveDocument.Masters.AddEx(visTypeDataGraphic) 
    Set vsoMasterCopy = vsoMaster.Open 
    Set vsoMaster_Old = ActiveDocument.Masters("old_master_name ") 
    Set vsoGraphicItem_Old = vsoMaster_Old.GraphicItems(1) 
    Set vsoGraphicItem = vsoMasterCopy.GraphicItems.AddCopy(vsoGraphicItem_Old) 
    vsoMasterCopy.Close     
 
End Sub
```


