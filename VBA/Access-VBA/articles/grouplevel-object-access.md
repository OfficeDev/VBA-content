---
title: GroupLevel Object (Access)
keywords: vbaac10.chm12247
f1_keywords:
- vbaac10.chm12247
ms.prod: access
api_name:
- Access.GroupLevel
ms.assetid: fdc4f24e-98aa-27bd-7a9d-271d48912dfa
ms.date: 06/08/2017
---


# GroupLevel Object (Access)

You can use the  **GroupLevel** property in Visual Basic to refer to the group level you are grouping or sorting on in a report.


## Remarks

The  **GroupLevel** property setting is an array in which each entry identifies a group level. To refer to a group level, use this syntax:

 **GroupLevel** ( _n_)

The number  _n_ is the group level, starting with 0. The first field or expression you group on is group level 0, the second is group level 1, and so on. You can have up to 10 group levels (0 to 9).

The following sample settings show how you use the  **GroupLevel** property to refer to a group level.



|**Group level**|**Refers to**|
|:-----|:-----|
|**GroupLevel** (0)|The first field or expression you sort or group on.|
|**GroupLevel** (1)|The second field or expression you sort or group on.|
|**GroupLevel** (2)|The third field or expression you sort or group on.|
You can use this property only by using Visual Basic to set the  **SortOrder**, **GroupOn**, **GroupInterval**, **KeepTogether**, and **ControlSource** properties. You set these properties in the **Open** event procedure of a report.

In reports, you can group or sort on more than one field or expression. Each field or expression you group or sort on is a group level.

You specify the fields and expressions to sort and group on by using the  **CreateGroupLevel** method.

If a group is already defined for a report (the  **GroupLevel** property is set to 0), then you can use the **ControlSource** property to change the group level in the report's Open event procedure. For example, the following code changes the **ControlSource** property to a value contained in the `txtPromptYou`text box on the open form named  `SortForm`:




```
Private Sub Report_Open(Cancel As Integer) 
 Me.GroupLevel(0).ControlSource _ 
 = Forms!SortForm!txtPromptYou 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](grouplevel-application-property-access.md)|
|[ControlSource](grouplevel-controlsource-property-access.md)|
|[GroupFooter](grouplevel-groupfooter-property-access.md)|
|[GroupHeader](grouplevel-groupheader-property-access.md)|
|[GroupInterval](grouplevel-groupinterval-property-access.md)|
|[GroupOn](grouplevel-groupon-property-access.md)|
|[KeepTogether](grouplevel-keeptogether-property-access.md)|
|[Parent](grouplevel-parent-property-access.md)|
|[Properties](grouplevel-properties-property-access.md)|
|[SortOrder](grouplevel-sortorder-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
