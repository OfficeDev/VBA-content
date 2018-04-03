---
title: Calculate Elapsed Time
ms.prod: access
ms.assetid: 90e46152-6d97-0860-a414-a17cc8ba40cf
ms.date: 06/08/2017
---


# Calculate Elapsed Time

This topic explains how Access stores the Date/Time data type and why you may receive unexpected results when you calculate or compare dates and times. 


## Storing Date/Time Data

Access stores the Date/Time data type as a double-precision, floating-point number (up to 15 decimal places). The integer portion of the double-precision number represents the date; the decimal portion represents the time. 

Valid date values range from -647,434 (January 1, 100 A.D.) to 2,958,465 (December 31, 9999 A.D.). A date value of 0 represents December 30, 1899. Access stores dates prior to December 30, 1899 as negative numbers. 

Valid time values range from .0 (00:00:00) to .99999 (23:59:59). The numeric value represents a fraction of one day. You can convert the numeric value into hours, minutes, and seconds by multiplying the numeric value by 24.


||||||
|:-----|:-----|:-----|:-----|:-----|
|**Double Number**|**Date Portion**|**Actual Date**|**Time Portion**|**Actual Time**|
|1.0|1|December 31,1899 |.0|12:00:00 A.M.|
|2.5|2|January 1, 1900 |.5 |12:00:00 P.M.|
|27468.96875|27468|March 15, 1975 |.96875 |11:15:00 P.M.|
|33914.125 |33914|November 6, 1992|||

## Calculating Time Data

Because a time value is stored as a fraction of a 24-hour day, you may receive incorrect formatting results when you calculate time intervals greater than 24 hours. To work around this behavior, you can create a user-defined function to ensure that time intervals are formatted correctly. 

The following procedure illustrates how to use the  **Format** function to format time intervals. The procedure accepts two time values and prints their the interval between them to the Immediate window in several different formats.




```vb
Function ElapsedTime(endTime As Date, startTime As Date) 
    Dim strOutput As String 
    Dim Interval As Date 
     
    ' Calculate the time interval. 
    Interval = endTime - startTime 
  
    ' Format and print the time interval in seconds. 
    strOutput = Int(CSng(Interval * 24 * 3600)) &; " Seconds" 
    Debug.Print strOutput 
         
    ' Format and print the time interval in minutes and seconds. 
    strOutput = Int(CSng(Interval * 24 * 60)) &; ":" &; Format(Interval, "ss") _ 
        &; " Minutes:Seconds" 
    Debug.Print strOutput 
     
    ' Format and print the time interval in hours, minutes and seconds. 
    strOutput = Int(CSng(Interval * 24)) &; ":" &; Format(Interval, "nn:ss") _ 
           &; " Hours:Minutes:Seconds" 
    Debug.Print strOutput 
         
    ' Format and print the time interval in days, hours, minutes and seconds. 
    strOutput = Int(CSng(Interval)) &; " days " &; Format(Interval, "hh") _ 
        &; " Hours " &; Format(Interval, "nn") &; " Minutes " &; _ 
        Format(Interval, "ss") &; " Seconds" 
    Debug.Print strOutput 
 
End Function
```


## Comparing Date Data

Because dates and times are stored together as double-precision numbers, you may receive unexpected results when you compare Date/Time data. For example, if you type the following expression in the Immediate window , you receive a  **False** (0) result even if today's date is 7/11/2006:


```vb
? Now()=DateValue("7/11/2006")
```

The  **[Now](http://msdn.microsoft.com/library/8F324994-2518-0C83-76C7-22CD67033B36%28Office.15%29.aspx)** function returns a double-precision number representing the current date and time. However, the **[DateValue](http://msdn.microsoft.com/library/8C9BD3D6-1614-EEB0-0714-4730EEEB1B95%28Office.15%29.aspx)** function returns an integer number representing the date but not a fractional time value. As a result, **Now** equals **DateValue** only when **Now** returns a time of 00:00:00 (12:00:00 A.M.).

To receive accurate results when you compare date values, use one of the following functions. To test each function, type it in the Immediate window, substitute the current date for 7/11/2006, and then press ENTER: 

To return an integer value, use the  **[Date](http://msdn.microsoft.com/library/8AFD02C8-C5B5-F8F3-FF8E-9A2AC0EA94B9%28Office.15%29.aspx)** function:




```vb
?Date()=DateValue("7/11/2006")
```

To remove the fractional portion of the  **Now** function, use the **[Int](http://msdn.microsoft.com/library/32CE40AC-FDF8-BD6D-E7F9-154C480A9602%28Office.15%29.aspx)** function:




```vb
?Int(Now())=DateValue("7/11/2006")
```


## Comparing Time Data

When you compare time values, you may receive inconsistent results because a time value is stored as the fractional portion of a double-precision, floating-point number. For example, if you type the following expression in the Immediate window, you receive a  **False** (0) result even though the two time values look the same:


```
var1 = #2:01:00 PM# 
var2 = DateAdd("n", 10, var1) 
? var2 = #2:11:00 PM# 
```

When Access converts a time value to a fraction, the calculated result may not be the exact equivalent of the time value. The small difference caused by the calculation is enough to produce a  **False** (0) result when you compare a stored value to a constant value.

To receive accurate results when you compare time values, use one of the following methods. To test each method, type it in the Immediate window, and then press ENTER: 

Add an associated date to the time comparison:




```
var1 = #7/11/2006 2:00:00 PM# 
var2 = DateAdd("n", 10, var1) 
? var2 = #7/11/2006 2:10:00 PM#
```

Convert the time values to  **String** data types before you compare them:




```
var1 = #2:00:00 PM# 
var2 = DateAdd("n", 10, var1) 
? CStr(var2) = CStr(#2:10:00 PM#)
```

Use the  **[DateDiff](http://msdn.microsoft.com/library/15C9DF5F-1403-B6A5-71B9-611E9820D804%28Office.15%29.aspx)** function to compare precise units such as seconds:




```
var1 = #2:00:00 PM# 
var2 = DateAdd("n", 10, var1) 
? DateDiff("s", var2, #2:10:00 PM#) = 0
```

 **Link provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community


- [Summing elapsed time that could go over 24 hours](http://www.utteraccess.com/wiki/index.php/Summing_elapsed_time_that_could_go_over_24_hours)
    

## About the Contributors
<a name="AboutContributors"> </a>

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 


