---
title: TimeScaleValues Object (Project)
ms.prod: project-server
ms.assetid: d94a0346-7cf5-b734-b32d-430fba980824
ms.date: 06/08/2017
---


# TimeScaleValues Object (Project)

Contains a collection of  **[TimeScaleValue](timescalevalue-object-project.md)** objects.


## Examples

 **Using the TimeScaleValue Object**

Use  **TimeScaleValues** ( _Index_ ), where _Index_ is the index number of the timescaled data item, to return a single **TimeScaleValue** object. The following example displays the number of hours of work per day for a resource during the first full week in October 2012.




```
Dim TSV As TimeScaleValues, HowMany As Long
Dim HoursPerDay As String

Set TSV = ActiveCell.Resource.TimeScaleData("10/1/2012", "10/5/2012", TimescaleUnit:=pjTimescaleDays)

For HowMany = 1 To TSV.Count
    HoursPerDay = HoursPerDay &amp; TSV(HowMany).StartDate &amp; " - " &amp; _
        TSV(HowMany).EndDate &amp; ", " &amp; TSV(HowMany) / 60 &amp; vbCrLf
Next HowMany

MsgBox HoursPerDay
```

 **Using the TimeScaleValues Collection**

Use the  **[TimeScaleData](http://msdn.microsoft.com/library/51649bc3-8224-15cd-dc9b-af37a1cc4d8b%28Office.15%29.aspx)** method to return a **TimeScaleValues** collection. The following example returns a **TimeScaleValues** collection for the amount of work done by the resource in the active cell between the specified dates, split into week-long portions.




```
ActiveCell.Resource.TimeScaleData("10/1/2012", "10/31/2012")
```

Use the  **[Add](http://msdn.microsoft.com/library/083ef154-31ce-55ec-793a-0627c1eff211%28Office.15%29.aspx)** method to add a **TimeScaleValue** object to the **TimeScaleValues** collection. The following example adds 8 hours of work to Tuesday of that week.




```
Dim TSV As TimeScaleValues
  
Set TSV = ActiveCell.Resource.TimeScaleData("10/1/2012", "10/5/2012", TimescaleUnit:=pjTimescaleDays)
TSV.Add 480, 2
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/083ef154-31ce-55ec-793a-0627c1eff211%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/58c5a8ae-0646-2f47-ad79-687ec8d41d4e%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/76ac63bf-74e1-3f1c-1089-90eb101e1147%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/8bbd6389-53ac-9f03-d155-c53e6a3dc681%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/1235dcdf-1cb0-23d3-f943-4e7acf513b40%28Office.15%29.aspx)|

## See also


#### Other resources


[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
