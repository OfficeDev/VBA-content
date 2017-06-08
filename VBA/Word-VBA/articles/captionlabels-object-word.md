---
title: CaptionLabels Object (Word)
ms.prod: word
ms.assetid: 7d18c0d6-6d58-9841-4665-ab13e2e2ad9f
ms.date: 06/08/2017
---


# CaptionLabels Object (Word)

A collection of  **[CaptionLabel](captionlabel-object-word.md)** objects that represent the available caption labels. The items in the **CaptionLabels** collection are listed in the **Label** box in the **Caption** dialog box.


## Remarks

Use the  **CaptionLabels** property to return the **CaptionLabels** collection. By default, the **CaptionLabels** collection includes the three built-in caption labels: Figure, Table, and Equation.

Use the  **[Add](captionlabels-add-method-word.md)** method to add a custom caption label. The following example adds a caption label named "Photo."




```
CaptionLabels.Add Name:="Photo"
```

Use  **CaptionLabels** (index), where index is the caption label name or index number, to return a single **CaptionLabel** object. The following example sets the numbering style for the Figure caption label.




```
CaptionLabels("Figure").NumberStyle = _ 
 wdCaptionNumberStyleLowercaseLetter
```

The index number represents the position of the caption label in the  **CaptionLabels** collection. The following example displays the first caption label.




```
MsgBox CaptionLabels(1).Name
```


## Methods



|**Name**|
|:-----|
|[Add](captionlabels-add-method-word.md)|
|[Item](captionlabels-item-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](captionlabels-application-property-word.md)|
|[Count](captionlabels-count-property-word.md)|
|[Creator](captionlabels-creator-property-word.md)|
|[Parent](captionlabels-parent-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
