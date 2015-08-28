
# IRibbonControl.Tag Property (Office)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Used to store arbitrary strings and fetch them at runtime. Read-only


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Tag**

 _expression_An expression that returns a  **IRibbonControl** object.


### Return Value

String


## Remarks
<a name="sectionSection1"> </a>

Normally you can distinguish between controls in a Ribbon user interface XML customization file using the  **Id** property. However, there are restrictions on what IDs can contain (no non-alphanumeric characters, and they must all be unique). The **Tag** property doesn't have these restrictions and so it can be used in the following situations, where ID doesn't work:


- If you need to store a special string with your control such as a filename. For example: tag="C:\path\file.xlsm."
    
- If you want multiple controls to be treated the same way by your callback procedures, but you don't want to maintain a list of all of their IDs (which must be unique). For example, you could have buttons on different tabs on the Ribbon, all with tag="blue", and then just check the  **Tag** property instead of the **ID** property when perfroming some common action.
    



## Example
<a name="sectionSection2"> </a>

In the XML used to customize the Ribbon user interface, you can set a tag as follows. When the MyFunction action is called, you can read the  **Tag** property, which will be equal to "some string".


```XML
<button id="mybutton" tag="some string" onAction="MyFunction"/>
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [IRibbonControl Object](63aef709-e1d3-b1a6-76af-b568ad0e69ae.md)
#### Other resources


 [IRibbonControl Object Members](396d85dc-ddd5-8985-0830-22ee5b1579dc.md)
