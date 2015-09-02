
# CapitalizeNamesOfDays Property

 **Last modified:** July 28, 2015

True if the first letter of day names is capitalized automatically. Read/write Boolean.

 _expression_. **CapitalizeNamesOfDays**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Example

This example sets Microsoft Graph to capitalize the first letter of the names of days.


```
With myChart.Application.AutoCorrect 
 .CapitalizeNamesOfDays = True 
 .ReplaceText = True 
End With
```

