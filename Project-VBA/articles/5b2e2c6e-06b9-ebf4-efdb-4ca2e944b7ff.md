
# StartDriver.EffectiveDateAdd Property (Project)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Gets the date and time that follows another date by a specified duration, using the effective calendar for a manually scheduled task. Read-only  **Variant**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **EffectiveDateAdd**( **_Date_**,  **_Duration_**)

 _expression_An expression that returns a  **StartDriver** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Date|Required| **Variant**|Arbitrary date and time, for example, "7/10/2010" or "7/10/2010 2:00:00 PM".|
|Duration|Required| **Variant**|Duration to add, for example, "3d" or "2w".|

## Remarks
<a name="sectionSection1"> </a>

The  **EffectiveDateAdd** property uses the effective calendar for manually scheduled tasks, which allows tasks to start and finish on non-working times. The property and arguments have no effect on actual task dates.

You can use the  ** [EffectiveDateSubtract](14529bd1-9029-d1bc-60a0-b7863cba4d6d.md)**,  **EffectiveDateAdd**, and  ** [EffectiveDateDifference](9b825839-31de-71f8-9804-015dfd5a293c.md)** properties to calculate start and finish dates for manually scheduled tasks.

To calculate a date for an automatically scheduled task, where you can also specify the calendar, use the  ** [DateAdd](df0da054-495c-c224-ebc8-b47acb78e2af.md)** method.


## Example
<a name="sectionSection2"> </a>

The following statement returns the value "7/9/2009 5:00:00 PM", which is six days after the specified date. 


```
Debug.Print ActiveProject.Tasks(3).StartDriver.EffectiveDateAdd("7/2/2009", "6d")
```

