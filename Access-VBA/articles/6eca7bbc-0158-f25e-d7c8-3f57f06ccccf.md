
# ChartTitle Object

 **Last modified:** July 28, 2015

Represents the title of the specified chart.

## Using the ChartTitle Object

Use the  **ChartTitle** property to return the **ChartTitle** object. The following example adds a title to the chart.


```
With myChart 
 .HasTitle = True 
 .ChartTitle.Text = "February Sales" 
End With
```


## Remarks

The  **ChartTitle** object doesn't exist and cannot be used unless the ** [HasTitle](9ecc48d3-fd86-e185-a69f-0676230b3194.md)**property for the chart is  **True**.

