
# OlkTextBox.MouseMove Event (Outlook)

 **Last modified:** July 28, 2015

Occurs after a mouse movement has been registered over the control.

## Syntax

 _expression_. **MouseMove**( **_Button_**,  **_Shift_**,  **_X_**,  **_Y_**)

 _expression_A variable that represents an  **OlkTextBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Button|Required| **Integer**|An  ** [OlMouseButton](f654f074-f7e7-6128-9d7d-8ec6adbfe5f7.md)** constant that specifies which button on the mouse has been pressed.|
|Shift|Required| **Integer**|A bitwise-OR mask of constants in the  ** [OlShiftState](f71dd27d-6930-1450-e8e9-11ab1eace6ca.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|
|X|Required| **[OLE_XPOS_CONTAINER]**|Identifies the location of the mouse cursor on the X-axis relative to the form.|
|Y|Required| **[OLE_YPOS_CONTAINER]**|Identifies the location of the mouse cursor on the Y-axis relative to the form.|

## Remarks

Pressing the  **ALT** key fires the **MouseMove** event.


## See also


#### Concepts


 [OlkTextBox Object](8c9438bf-e20a-2f70-90ac-097cf09594ca.md)
#### Other resources


 [OlkTextBox Object Members](f4a5f9ea-15f7-164e-d7ca-77a0842105c8.md)
