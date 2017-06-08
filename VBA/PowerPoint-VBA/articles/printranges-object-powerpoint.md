---
title: PrintRanges Object (PowerPoint)
keywords: vbapp10.chm518000
f1_keywords:
- vbapp10.chm518000
ms.prod: powerpoint
api_name:
- PowerPoint.PrintRanges
ms.assetid: 5c1e9dc1-e30c-bc65-5283-448b95795b11
ms.date: 06/08/2017
---


# PrintRanges Object (PowerPoint)

A collection of all the  **[PrintRange](printrange-object-powerpoint.md)** objects in the specified presentation. Each **PrintRange** object represents a range of consecutive slides or pages to be printed.


## Example

Use the [Ranges](printoptions-ranges-property-powerpoint.md)property to return the  **PrintRanges** collection. The following example clears all previously defined print ranges from the collection for the active presentation.


```vb
ActivePresentation.PrintOptions.Ranges.ClearAll
```

Use the [Add](printranges-add-method-powerpoint.md)method to create a  **PrintRange** object and add it to the **PrintRanges** collection. The following example defines three print ranges that represent slide 1, slides 3 through 5, and slides 8 and 9 in the active presentation and then prints the slides in these ranges.




```vb
With ActivePresentation.PrintOptions

    .RangeType = ppPrintSlideRange

    With .Ranges

        .ClearAll

        .Add 1, 1

        .Add 3, 5

        .Add 8, 9

    End With

End With

ActivePresentation.PrintOut
```

Use  **Ranges** (index), where index is the print range index number, to return a single **PrintRange** object. The following example displays a message that indicates the starting and ending slide numbers for print range one in the active presentation.




```vb
With ActivePresentation.PrintOptions.Ranges
    If .Count > 0 Then
        With .Item(1)
            MsgBox "Print range 1 starts on slide " &; .Start &; _
                " and ends on slide " &; .End
        End With
    End If
End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

