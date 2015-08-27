
# VisToParts Enumeration (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Values returned by the  **Connect.ToPart** property.


## Remarks

The  **VisToParts** return codes indicate the part of a shape to which a connection is made.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visConnectionPoint**|100 + row index of connection point|Connect to specified connection point on target shape.|
| **visConnectToError**|-1|Error connecting to shape.|
| **visGuideIntersect**|4|Connect to intersection of guides on target shape.|
| **visGuideX**|1|Connect to vertical guide on target shape.|
| **visGuideY**|2|Connect to horizontal guide on target shape.|
| **visToAngle**|7|Connect to angle on target shape.|
| **visToNone**|0|Do not connect.|
| **visWholeShape**|3|Connect to entire target shape, using dynamic glue.|
