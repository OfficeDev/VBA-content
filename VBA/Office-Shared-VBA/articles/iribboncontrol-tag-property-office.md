---
title: IRibbonControl.Tag Property (Office)
keywords: vbaof11.chm288003
f1_keywords:
- vbaof11.chm288003
ms.prod: office
api_name:
- Office.IRibbonControl.Tag
ms.assetid: d0f041c0-d7bc-7a4f-df9b-ba62fa08f1ca
ms.date: 06/08/2017
---


# IRibbonControl.Tag Property (Office)

Used to store arbitrary strings and fetch them at runtime. Read-only


## Syntax

 _expression_. **Tag**

 _expression_ An expression that returns a **IRibbonControl** object.


### Return Value

String


## Remarks

Normally you can distinguish between controls in a Ribbon user interface XML customization file using the  **Id** property. However, there are restrictions on what IDs can contain (no non-alphanumeric characters, and they must all be unique). The **Tag** property doesn't have these restrictions and so it can be used in the following situations, where ID doesn't work:


- If you need to store a special string with your control such as a filename. For example: tag="C:\path\file.xlsm."
    
- If you want multiple controls to be treated the same way by your callback procedures, but you don't want to maintain a list of all of their IDs (which must be unique). For example, you could have buttons on different tabs on the Ribbon, all with tag="blue", and then just check the  **Tag** property instead of the **ID** property when perfroming some common action.
    



## Example

In the XML used to customize the Ribbon user interface, you can set a tag as follows. When the MyFunction action is called, you can read the  **Tag** property, which will be equal to "some string".


```XML
<button id="mybutton" tag="some string" onAction="MyFunction"/>
```


## See also


#### Concepts


[IRibbonControl Object](iribboncontrol-object-office.md)
#### Other resources


[IRibbonControl Object Members](iribboncontrol-members-office.md)

