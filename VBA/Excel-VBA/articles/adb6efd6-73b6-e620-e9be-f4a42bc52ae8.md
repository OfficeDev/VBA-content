
# Borders Object (Excel)

A collection of four  **[Border](bca516bf-7c0f-f9df-078d-dfb522f256f3.md)** objects that represent the four borders of a **[Range](b8207778-0dcc-4570-1234-f130532cc8cd.md)** or **[Style](3c1e9184-0075-5f46-9a1a-0b61d874d1f8.md)** object.


## Remarks

Use the  **Borders** property to return the **Borders** collection, which contains all four borders.

You can set border properties for an individual border only with  **Range** and **Style** objects. Other bordered objects, such as error bars and series lines, have a border that's treated as a single entity, regardless of how many sides it has. For these objects, you must return and set properties for the entire border as a unit. For more information, see the **Border** object.


## Example

The following example adds a double border to cell A1 on worksheet one.


```
Worksheets(1).Range("A1").Borders.LineStyle = xlDouble
```

Use  **Borders** ( _index_ ), where _index_ identifies the border, to return a single **Border** object. The following example sets the color of the bottom border of cells A1:G1 to red.




```
Worksheets("Sheet1").Range("A1:G1"). _ 
 Borders(xlEdgeBottom).Color = RGB(255, 0, 0)
```

 _Index_ can be one of the following **[xlBordersIndex](91ab77e7-c54f-266d-fc61-7ce0bed1bd8c.md)** constants: **xlDiagonalDown**, **xlDiagonalUp**, **xlEdgeBottom**, **xlEdgeLeft**, **xlEdgeRight**, or **xlEdgeTop**, **xlInsideHorizontal**, or **xlInsideVertical**.


## Properties



|**Name**|
|:-----|
|[Application](bba16d88-5609-3792-3ace-9928fdaccd98.md)|
|[Color](3ee1bce3-56e2-c93f-432f-8f1d037a7624.md)|
|[ColorIndex](fe0a7b5e-254d-c773-88cc-70728db44840.md)|
|[Count](fe015e4c-89f3-cb8c-5215-55181dcdc0c4.md)|
|[Creator](00a52b71-0faa-e52c-ad65-f33e684187f9.md)|
|[Item](19184379-d551-396e-8cb6-ff240e3c85fa.md)|
|[LineStyle](a057234d-0442-3fd7-5547-b19451774c0e.md)|
|[Parent](43a8a82d-d2b9-59d3-36b2-97ffffea6cdb.md)|
|[ThemeColor](ca1d3f82-af14-f5be-71f3-3ba0c340ebbf.md)|
|[TintAndShade](29c591bf-311e-5706-0222-1db144a92b77.md)|
|[Value](9415589c-f698-a09d-d232-cf2ca32e6b11.md)|
|[Weight](cdf2d0d2-9c4d-1b07-38fc-3828126c77bf.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)