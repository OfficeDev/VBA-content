
# Year Function (Access custom web app)
Returns a numeric value that represents the year of the specified date in the Gregorian calendar.

 **Last modified:** March 09, 2015

 _ **Applies to:** Access 2013 | Access 2016_

## Syntax

 **Year** ( _Date_ )

The  **Year** function contains the following arguments.



|**Argument name**|**Description**|
|:-----|:-----|
| _Date_|An expression that can be resolved to a Date/Time value. The  _Date_ argument expression, column expression, user-defined variable or string literal.|

## Remarks

Values returned by the  **Year**, **Month**, and **Day** functions will be Gregorian values regardless of the display format for the supplied date value. For example, if the display format of the supplied date uses the Hijri calendar, the returned values for the **Year**, **Month**, and **Day** functions will be values associated with the equivalent Gregorian date.

