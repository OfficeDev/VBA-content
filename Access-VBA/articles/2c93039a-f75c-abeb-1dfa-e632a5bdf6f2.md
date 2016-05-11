
# DisplayHourglassPointer Macro Action

 **Last modified:** July 28, 2015

 _ **Applies to:** Access 2013 | Access 2016_

You can use the  **DisplayHourglassPointer** action to change the mouse pointer to an image of an hourglass (or another icon you've chosen) while a macro is running. This action can provide a visual indication that the macro is running. This is especially useful when a macro action or the macro itself takes a long time to run.


## Setting

The  **DisplayHourglassPointer** action has the following argument.



|**Action argument**|**Description**|
|:-----|:-----|
|**Hourglass On**|Click  **Yes** (display the icon) or **No** (display the normal mouse pointer) in the **Hourglass On** box in the **Action Arguments** section of the Macro Builder pane. The default is **Yes**.|

## Remarks

You often use this action if you have turned echo off by using the  **Echo** action. When echo is off, Access suspends screen updates until the macro is finished.

Access automatically resets the  **Hourglass On** argument to **No** when the macro finishes running.


 **Note**  

To run the  **DisplayHourglassPointer** action in a Visual Basic for Applications (VBA) module, use the **Hourglass** method of the **DoCmd** object.

