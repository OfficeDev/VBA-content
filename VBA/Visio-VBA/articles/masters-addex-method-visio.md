---
title: Masters.AddEx Method (Visio)
keywords: vis_sdr.chm10851450
f1_keywords:
- vis_sdr.chm10851450
ms.prod: visio
api_name:
- Visio.Masters.AddEx
ms.assetid: a27b1a7c-37f4-90c9-91f1-2249611a3cf9
ms.date: 06/08/2017
---


# Masters.AddEx Method (Visio)

Adds a new  **Master** object of the specified type to the **Masters** collection of a Microsoft Visio document.


## Syntax

 _expression_ . **AddEx**( **_Type_** )

 _expression_ A variable that represents a **Masters** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **VisMasterTypes**|A master type from the  **VisMasterTypes** enumeration. See Remarks for possible values.|

### Return Value

Master


## Remarks

For the  _Type_ parameter, pass one of the following members of **VisMasterTypes** , which is declared in the Visio type library.



|**Constant**|**Value **|**Description**|
|:-----|:-----|:-----|
| **visTypeMaster**|1|Creates a shape master.|
| **visTypeFillPattern**|2|Creates a fill-pattern master.|
| **visTypeLinePattern**|3|Creates a line-pattern master.|
| **visTypeLineEnd**|4|Creates a line-end master.|
| **visTypeDataGraphic**|5|Creates a data graphic master.|
| **visTypeThemeColors**|6|Creates a theme-colors master.|
| **visTypeThemeEffects**|7|Creates a theme-effects master.|
The  **AddEx** method returns the **Master** object added.

If the master added is of type  **visTypeDataGraphic** , Visio names it "Data Graphic", and if it is not the first data graphic in the **Masters** collection of the document, Visio appends the index number of the master in the collection to the name. For example, if there were already 5 objects in the **Masters** collection, one of which was a data graphic, the next data graphic added would be named "Data Graphic.6".

Naming of masters of type  **visTypeThemeColors** and **visTypeThemeEffects** follows the same pattern, and the resulting new masters are named "Theme Colors. _x_ " and "Theme Effects. _x_ " respectively, where _x_ is the index number in the collection. Masters of all other types are simply named "Master. _x_ ".


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **AddEx** method to add a new **Master** object of type **visTypeDataGraphic** to the **Masters** collection of the active document.


```vb
Public Sub AddEx_Example() 
 
    Dim vsoMaster As Visio.Master 
     
    Set vsoMaster = Visio.ActiveDocument.Masters.AddEx(visTypeDataGraphic) 
     
    Debug.Print vsoMaster.Name 
 
End Sub
```


