---
title: OLEFormat.DoVerb Method (PowerPoint)
keywords: vbapp10.chm562007
f1_keywords:
- vbapp10.chm562007
ms.prod: powerpoint
api_name:
- PowerPoint.OLEFormat.DoVerb
ms.assetid: 1ee39c5d-3646-81de-79e9-f8cff869308d
ms.date: 06/08/2017
---


# OLEFormat.DoVerb Method (PowerPoint)

Requests that an OLE object perform one of its verbs. 


## Syntax

 _expression_. **DoVerb**( **_Index_** )

 _expression_ A variable that represents an **OLEFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional|**Integer**|The verb to perform. If this argument is omitted, the default verb is performed.|

## Remarks

Use the  **[ObjectVerbs](oleformat-objectverbs-property-powerpoint.md)** property to determine the available verbs for an OLE object.


## Example

This example performs the default verb for shape three on slide one in the active presentation if shape three is a linked or embedded OLE object.


```vb
With ActivePresentation.Slides(1).Shapes(3)
    If .Type = msoEmbeddedOLEObject Or _
            .Type = msoLinkedOLEObject Then
        .OLEFormat.DoVerb
    End If
End With
```

This example performs the verb "Open" for shape three on slide one in the active presentation if shape three is an OLE object that supports the verb "Open."




```vb
With ActivePresentation.Slides(1).Shapes(3)
    If .Type = msoEmbeddedOLEObject Or _
            .Type = msoLinkedOLEObject Then

        For Each sVerb In .OLEFormat.ObjectVerbs
            nCount = nCount + 1
            If sVerb = "Open" Then
                .OLEFormat.DoVerb nCount
                Exit For
            End If
        Next
    End If
End With
```


## See also


#### Concepts


[OLEFormat Object](oleformat-object-powerpoint.md)

