---
title: Tags.Name Method (PowerPoint)
keywords: vbapp10.chm611008
f1_keywords:
- vbapp10.chm611008
ms.prod: powerpoint
api_name:
- PowerPoint.Tags.Name
ms.assetid: da2317a5-927b-82c9-2ebe-408c1dcd1e06
ms.date: 06/08/2017
---


# Tags.Name Method (PowerPoint)

Returns the name of the specified tag as a  **String**.


## Syntax

 _expression_. **Name**( **_Index_** )

 _expression_ A variable that represents a **Tags** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|The tag number.|

### Return Value

String


## Example

This example displays the name and value for each tag associated with slide one in the active presentation.


```vb
With Application.ActivePresentation.Slides(1).Tags

    For i = 1 To .Count

        MsgBox "Tag #" &; i &; ": Name = " &; .Name(i)

        MsgBox "Tag #" &; i &; ": Value = " &; .Value(i)

    Next

End With
```

This example searches through the tags for each slide in the active presentation. If there is a tag named "PRIORITY," a message box displays the tag value. If there is no tag named "PRIORITY," the example adds this tag, giving it the value "Unknown."




```vb
For Each s In Application.ActivePresentation.Slides
    With s.Tags
        found = False
        For i = 1 To .Count
            If .Name(i) = "PRIORITY" Then
                found = True
                slNum = .Parent.SlideIndex
                MsgBox "Slide " &; slNum &; _
                    " priority: " &; .Value(i)
            End If
        Next

        If Not found Then
            slNum = .Parent.SlideIndex
            .Add "Name", "New Figures"
            .Add "Priority", "Unknown"
            MsgBox "Slide " &; slNum &; _
                " priority tag added: Unknown"
        End If
    End With
Next
```


## See also


#### Concepts


[Tags Object](tags-object-powerpoint.md)

