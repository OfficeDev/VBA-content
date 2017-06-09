---
title: Presentations Object (PowerPoint)
keywords: vbapp10.chm522000
f1_keywords:
- vbapp10.chm522000
ms.prod: powerpoint
api_name:
- PowerPoint.Presentations
ms.assetid: 0b952edc-8628-71ef-e854-3bcefbb3bc61
ms.date: 06/08/2017
---


# Presentations Object (PowerPoint)

A collection of all the  **[Presentation](presentation-object-powerpoint.md)** objects in Microsoft PowerPoint. Each **Presentation** object represents a presentation that's currently open in PowerPoint.


## Remarks

The  **Presentations** collection doesn't include open add-ins, which are a special kind of hidden presentation. You can, however, return a single open add-in if you know its file name. For example `Presentations("oscar.ppa")` will return the open add-in named "Oscar.ppa" as a **Presentation** object. However, it is recommended that the **AddIns** collection be used to return open add-ins.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.PowerPoint** reference, this collection maps to the following types:


-  **Microsoft.Office.Interop.PowerPoint.Presentations.GetEnumerator** (to enumerate the **Presentation** objects.)
    

## Example

Use the [Presentations](http://msdn.microsoft.com/library/d6f5f565-d593-e230-c3b9-2302bdd83644%28Office.15%29.aspx) property to return the **Presentations** collection. Use the[Add](http://msdn.microsoft.com/library/9a09ad9b-c52d-9fd6-20ef-68b694596ed2%28Office.15%29.aspx) method to create a new presentation and add it to the collection. The following example creates a new presentation, adds a slide to the presentation, and then saves the presentation.


```
Set newPres = Presentations.Add(True) 
newPres.Slides.Add 1, 1 
newPres.SaveAs "Sample"
```

Use  **Presentations** (index), where index is the presentation's name or index number, to return a single **Presentation** object. The following example prints presentation one.




```
Presentations(1).PrintOut
```

Use the [Open](http://msdn.microsoft.com/library/c19456ba-e5a8-83da-00ae-dd387e38febf%28Office.15%29.aspx) method to open a presentation and add it to the **Presentations** collection. The following example opens the file Sales.ppt as a read-only presentation.




```
Presentations.Open FileName:="sales.ppt", ReadOnly:=True
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/9a09ad9b-c52d-9fd6-20ef-68b694596ed2%28Office.15%29.aspx)|
|[CanCheckOut](http://msdn.microsoft.com/library/60393f0c-11e1-169d-2ead-c6556f1d1364%28Office.15%29.aspx)|
|[CheckOut](http://msdn.microsoft.com/library/c6145ab1-f6d5-265a-8244-40b5fa67aedf%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/f0d84e16-4d94-dd74-9e6f-4e57edfdc72d%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/c19456ba-e5a8-83da-00ae-dd387e38febf%28Office.15%29.aspx)|
|[Open2007](http://msdn.microsoft.com/library/45bbbe1f-461c-d908-0d3b-8b4e8aa681a6%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/5c42ecee-19ce-6e00-9aed-556fe32daf8b%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/e9f4d85f-4ba3-6c07-353d-79bbf39f91da%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/5c1e9107-2b42-0b06-ddbc-6ed0186e96d2%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
