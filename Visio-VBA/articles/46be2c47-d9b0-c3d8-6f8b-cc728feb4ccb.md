
# Window.ShowScrollBars Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Determines whether scroll bars are shown in the drawing window. Read/write.


## Syntax

 _expression_. **ShowScrollBars**

 _expression_A variable that represents a  **Window** object.


### Return Value

Integer


## Remarks

The  **ShowScrollBars** property value can be any combination of the following **VisScrollbarStates** constants, which are declared in the Visio type library. The default value is **visScrollBarBoth**, which shows both horizontal and vertical scroll bars.



|**Constant **|**Value **|
|:-----|:-----|
| **visScrollBarNeither**|&amp;H0 |
| **visScrollBarHoriz**|&amp;H1 |
| **visScrollBarVert**|&amp;H4 |
| **visScrollBarBoth**|&amp;H5 |
