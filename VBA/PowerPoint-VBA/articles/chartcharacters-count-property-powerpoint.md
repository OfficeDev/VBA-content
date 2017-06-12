---
title: ChartCharacters.Count Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ChartCharacters.Count
ms.assetid: 99e1634b-49de-220e-e0e1-cfb31a1ba73a
ms.date: 06/08/2017
---


# ChartCharacters.Count Property (PowerPoint)

Returns the number of objects in the collection. Read-only  **Long**.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **[ChartCharacters](chartcharacters-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example makes the last character a superscript character in the title of the first chart in the active document.




```vb
Sub MakeSuperscript()

    Dim n As Integer



    With ActiveDocument.InlineShapes(1)

        If .HasChart Then

            n = .Chart.Title.Characters.Count

            .Chart.Title.Characters(n, 1).Font.Superscript = True

        End If

    End With

End Sub
```


## See also


#### Concepts


[ChartCharacters Object](chartcharacters-object-powerpoint.md)

