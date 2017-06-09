---
title: TimeScaleValue Object (Project)
ms.prod: project-server
api_name:
- Project.TimeScaleValue
ms.assetid: bea0ad82-a3de-30d8-f191-dc2248c32653
ms.date: 06/08/2017
---


# TimeScaleValue Object (Project)

Represents a timescaled data item. The  **TimeScaleValue** object is a member of the **[TimeScaleValues](timescalevalues-object-project.md)** collection.


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
|[Clear](http://msdn.microsoft.com/library/3ed3a584-5496-cdf4-eafa-e0ecdd01edfd%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/ebe03270-1713-77f9-1ac9-97922b2aa612%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/feab3c92-a313-9ff0-4549-69465f6a3e6f%28Office.15%29.aspx)|
|[EndDate](http://msdn.microsoft.com/library/e9acd4f8-b002-5195-2e0c-505b633a3b54%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/ebb523d2-cf85-180c-6808-ea83c8d8a5ba%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/69b3a11e-609a-5d10-a76c-5e524e75c453%28Office.15%29.aspx)|
|[StartDate](http://msdn.microsoft.com/library/fdd70c48-7f07-f4dc-db93-ad46fb30a2bb%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/30665b24-bc19-a6a2-cb1b-a70c3736b05b%28Office.15%29.aspx)|

