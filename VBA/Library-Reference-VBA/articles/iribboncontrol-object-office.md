---
title: IRibbonControl Object (Office)
keywords: vbaof11.chm288000
f1_keywords:
- vbaof11.chm288000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.IRibbonControl
ms.assetid: 63aef709-e1d3-b1a6-76af-b568ad0e69ae
---


# IRibbonControl Object (Office)

Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility.


## Remarks

The  **IRibbonControl** object contains the name (ID) of the control and the current **Window** object for the Ribbon UI control.


## Example

The following example, written in C#, shows two procedures called from the  **onAction** event procedure of a Button control and a ToggleButton control. In the first procedure, the **IRibbonControl** object representing the control is passed into the procedure and a message box is displayed indicating that the button was pressed along with the ID of the button. The second procedure is similar to the first with the addition of a **Boolean** parameter indicating that the button was pressed.


```
public void ButtonOnAction(IRibbonControl control) 
{ 
 MessageBox.Show("Button clicked: " + control.Id); 
} 
 
public void ToggleButtonOnAction(IRibbonControl control, bool pressed) 
{ 
...if (pressed) 
 MessageBox.Show("ToggleButton was switched on."); 
 else 
 MessageBox.Show("ToggleButton was switched off."); 
}
```


## Properties



|**Name**|
|:-----|
|[Context](http://msdn.microsoft.com/library/39f9d85a-00e9-9682-3957-51d9e72b4d83%28Office.15%29.aspx)|
|[Id](http://msdn.microsoft.com/library/56a0d143-66de-ab77-0c21-d34341ce5da4%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/d0f041c0-d7bc-7a4f-df9b-ba62fa08f1ca%28Office.15%29.aspx)|

## See also


#### Other resources


[IRibbonControl Object Members](http://msdn.microsoft.com/library/396d85dc-ddd5-8985-0830-22ee5b1579dc%28Office.15%29.aspx)
[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
