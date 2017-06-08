---
title: SmartArtNode.OrgChartLayout Property (Office)
ms.prod: office
api_name:
- Office.SmartArtNode.OrgChartLayout
ms.assetid: 183879a1-94fe-e102-51ec-66146d002f75
ms.date: 06/08/2017
---


# SmartArtNode.OrgChartLayout Property (Office)

Retrieves or sets the  **MsoOrgChartLayoutType** associated with this node if there is one. Read/write


## Syntax

 _expression_. **OrgChartLayout**

 _expression_ An expression that returns a **SmartArtNode** object.


## Remarks

Possible members are:


- msoOrgChartLayoutBothHanging
    
- msoOrgChartLayoutDefault
    
- msoOrgChartLayoutLeftHanging
    
- msoOrgChartLayoutMixed
    
- msoOrgChartLayoutRightHanging
    
- msoOrgChartLayoutStandard
    

## Example

The following code sets the OrgChartLayout property to the default layout.


```
Dim saNode As SmartArtNode 
saNode.OrgChartLayout = msoOrgChartLayoutDefault
```


## See also


#### Concepts


[SmartArtNode Object](smartartnode-object-office.md)
#### Other resources


[SmartArtNode Object Members](smartartnode-members-office.md)

