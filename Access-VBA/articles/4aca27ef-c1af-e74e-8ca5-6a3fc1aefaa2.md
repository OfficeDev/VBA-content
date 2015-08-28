
# MinimumScale Property

 **Last modified:** July 28, 2015

Returns or sets the minimum value on the axis. Read/write  **Double**.

## Remarks

Setting this property sets the  ** [MinimumScaleIsAuto](95ed7a2b-efda-b05a-da2e-789a166a97c8.md)**property to  **False**.


## Example

This example sets the minimum and maximum values for the value axis.


```
With myChart.Axes(xlValue) 
 .MinimumScale = 10 
 .MaximumScale = 120 
End With
```

