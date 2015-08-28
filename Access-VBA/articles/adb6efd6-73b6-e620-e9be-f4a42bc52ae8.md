
# Borders Object (Excel)

 **Last modified:** July 28, 2015

A collection of four  ** [Border](bca516bf-7c0f-f9df-078d-dfb522f256f3.md)** objects that represent the four borders of a ** [Range](b8207778-0dcc-4570-1234-f130532cc8cd.md)**or  ** [Style](3c1e9184-0075-5f46-9a1a-0b61d874d1f8.md)** object.

## Remarks

Use the  **Borders** property to return the **Borders** collection, which contains all four borders.

You can set border properties for an individual border only with  **Range** and **Style** objects. Other bordered objects, such as error bars and series lines, have a border that's treated as a single entity, regardless of how many sides it has. For these objects, you must return and set properties for the entire border as a unit. For more information, see the **Border** object.


## Example

The following example adds a double border to cell A1 on worksheet one.


```
Worksheets(1).Range("A1").Borders.LineStyle = xlDouble
```

Use  **Borders**( _index_), where  _index_ identifies the border, to return a single **Border** object. The following example sets the color of the bottom border of cells A1:G1 to red.




```
Worksheets("Sheet1").Range("A1:G1"). _ 
 Borders(xlEdgeBottom).Color = RGB(255, 0, 0)
```

 _Index_ can be one of the following ** [xlBordersIndex](91ab77e7-c54f-266d-fc61-7ce0bed1bd8c.md)** constants: **xlDiagonalDown**,  **xlDiagonalUp**,  **xlEdgeBottom**,  **xlEdgeLeft**,  **xlEdgeRight**, or  **xlEdgeTop**,  **xlInsideHorizontal**, or  **xlInsideVertical**.


## See also


#### Concepts


 [Excel Object Model Reference](11ea8598-8a20-92d5-f98b-0da04263bf2c.md)
#### Other resources


 [Borders Object Members](8fb1ee1d-8e09-0b65-a9a3-4f278f6f9164.md)
