---
title: Table Object (PowerPoint)
keywords: vbapp10.chm622000
f1_keywords:
- vbapp10.chm622000
ms.prod: powerpoint
api_name:
- PowerPoint.Table
ms.assetid: ebbbca9f-4591-10ce-3c74-33b46a3b7cdf
ms.date: 06/08/2017
---


# Table Object (PowerPoint)

Represents a table shape on a slide. The  **Table** object is a member of the **Shapes** collection. The **Table** object contains the **[Columns](http://msdn.microsoft.com/library/ba2fb830-bb60-b259-3a3f-1281f77d6368%28Office.15%29.aspx)** collection and the **[Rows](rows-object-powerpoint.md)** collection.


## Example

Use  **Shapes** (index), where index is a number, to return a shape containing a table. Use the[HasTable](http://msdn.microsoft.com/library/fa38891a-e915-3a5c-4169-3c14e5e7136e%28Office.15%29.aspx)property to see if a shape contains a table. This example walks through the shapes on slide one, checks to see if each shape has a table, and then sets the mouse click action for each table shape to advance to the next slide.


```
With ActivePresentation.Slides(2).Shapes

    For i = 1 To .Count

        If .Item(i).HasTable Then

            .Item(i).ActionSettings(ppMouseClick) _

                .Action = ppActionNextSlide

        End If

    Next

End With
```

Use the [Cell](http://msdn.microsoft.com/library/31a2908b-7a33-994d-860a-e01da62729e7%28Office.15%29.aspx)method of the  **Table** object to access the contents of each cell. This example inserts the text "Cell 1" in the first cell of the table in shape five on slide three.




```
ActivePresentation.Slides(3).Shapes(5).Table _

    .Cell(1, 1).Shape.TextFrame.TextRange _

    .Text = "Cell 1"
```

Use the [AddTable](http://msdn.microsoft.com/library/77ce193e-10f7-25f4-a6f8-99d7d2b781ad%28Office.15%29.aspx)method to add a table to a slide. This example adds a 3x3 table on slide two in the active presentation.




```
ActivePresentation.Slides(2).Shapes.AddTable(3, 3)
```


## Methods



|**Name**|
|:-----|
|[ApplyStyle](http://msdn.microsoft.com/library/3e03bee2-d066-8687-f0cb-3b2460f44bbf%28Office.15%29.aspx)|
|[Cell](http://msdn.microsoft.com/library/31a2908b-7a33-994d-860a-e01da62729e7%28Office.15%29.aspx)|
|[ScaleProportionally](http://msdn.microsoft.com/library/1c703fe7-d657-5588-1991-23304a5b2bda%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AlternativeText](http://msdn.microsoft.com/library/db35ce8c-0115-4e72-db25-3d555242aee4%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/7284f690-269f-f9fb-5898-99db1b47e5f2%28Office.15%29.aspx)|
|[Background](http://msdn.microsoft.com/library/160ff59b-fe7e-16dd-4cf5-21f270e56ffc%28Office.15%29.aspx)|
|[Columns](http://msdn.microsoft.com/library/0645fa19-d5a2-1f4c-ae15-9623925d39bc%28Office.15%29.aspx)|
|[FirstCol](http://msdn.microsoft.com/library/34eb7612-f3df-3cbb-4a51-911bdcd065ab%28Office.15%29.aspx)|
|[FirstRow](http://msdn.microsoft.com/library/49a38e0b-7f30-b89f-7ee1-e45d60c2270f%28Office.15%29.aspx)|
|[HorizBanding](http://msdn.microsoft.com/library/58d864a2-6a5e-2860-b656-f7dc06d05de0%28Office.15%29.aspx)|
|[LastCol](http://msdn.microsoft.com/library/cf06f919-4092-8a8e-3fed-74456a507dd9%28Office.15%29.aspx)|
|[LastRow](http://msdn.microsoft.com/library/b3cf6345-42bf-f371-3e70-f4d62b11f15d%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/1c873300-6a8d-bdd7-ff69-aa0ffc9aa157%28Office.15%29.aspx)|
|[Rows](http://msdn.microsoft.com/library/f7003d61-62d4-8d00-15c5-d9a2c5d57625%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/04a1e090-8d1e-95b8-2ea3-306db29be866%28Office.15%29.aspx)|
|[TableDirection](http://msdn.microsoft.com/library/3fbb1c4b-6cdb-f97e-7b85-c41897bc5ced%28Office.15%29.aspx)|
|[Title](http://msdn.microsoft.com/library/bbaf0307-22ce-d6d7-8996-ff7758bffab3%28Office.15%29.aspx)|
|[VertBanding](http://msdn.microsoft.com/library/dff08449-332d-34af-37e4-2e0a3edff069%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
