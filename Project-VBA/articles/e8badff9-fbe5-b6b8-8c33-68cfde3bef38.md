
# Shape.ZOrder Method (Project)
Moves the shape in front of or behind other shapes (that is, changes the position in the z-order).

 **Last modified:** July 28, 2015


## Syntax

 _expression_. **ZOrder**(ZOrderCmd)

 _expression_A variable that represents a  **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|ZOrderCmd|Required| ** [MsoZOrderCmd](http://msdn.microsoft.com/en-us/library/office/ff861432%28v=office.15%29)**|Specifies where to move the shape relative to the other shapes.|
|ZOrderCmd|Required|MSOZORDERCMD||

### Return value

 **Nothing**


## Remarks

Use the  **ZOrderPosition** property to determine the current position of a shape in the z-order.


## See also


#### Other resources


 [Shape Object](d2b32bcd-5595-a4a7-9772-feb25fd0103a.md)
 [MsoZOrderCmd](http://msdn.microsoft.com/en-us/library/office/ff861432%28v=office.15%29)
 [ZOrderPosition Property](d9f0d46f-65b1-bb1f-cb75-ce4d7c3b3ab2.md)
