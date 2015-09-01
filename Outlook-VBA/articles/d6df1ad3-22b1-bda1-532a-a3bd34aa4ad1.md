
# SelectNamesDialog.SetDefaultDisplayMode Method (Outlook)

 **Last modified:** July 28, 2015

Sets the default display mode for the  **Select Names** dialog box, specifying its caption and button labels.

## Syntax

 _expression_. **SetDefaultDisplayMode**( **_defaultMode_**)

 _expression_A variable that represents a  **SelectNamesDialog** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|defaultMode|Required| ** [OlDefaultSelectNamesDisplayMode](4a9c2183-c704-fc4d-e3c8-32c53b9688bb.md)**|A constant in the  **OlDefaultSelectNamesDisplayMode** enumeration that determines the default caption and button labels for the **Select Names** dialog box.|

## Remarks

 **SetDefaultDisplayMode** is optional. If you do not call **SetDefaultDisplayMode** before calling ** [Display](a689dfca-e4f7-f1c0-03a1-71e7d7e310b7.md)**, the default display mode will be  **OlDefaultSelectNamesDisplayMode.olDefaultMail**. To set the display mode to a different value, you should call  **SetDefaultDisplayMode** before calling the **Display** method.

This method allows you to display the  **Select Names** dialog box without using a resource file to localize the values for the caption, the **To** label, **Cc** label, and **Bcc** label. You can override the built-in behavior by setting your own values for ** [Caption](a728bcb5-8eee-8f77-76d7-4c15d53d79e2.md)**,  ** [ToLabel](1c2f15fd-57c6-e0a5-923c-2b3b217bb7a0.md)**,  ** [CcLabel](b28def6f-725c-ba65-cf7f-4abbc7ba3cb8.md)**, and  ** [BccLabel](9c826c3e-c7d3-6fd0-f900-24ba31925681.md)**. 

You can set additional properties (for example, setting  ** [NumberOfRecipientSelectors](2cb40e5f-b122-d032-9343-54fe98bc5455.md)** to **olRecipientSelectors.olToCc**) after calling  **SetDefaultDisplayMode**. The  **Select Names** dialog box will observe the subsequent setting.


## See also


#### Concepts


 [SelectNamesDialog Object](1522736a-3cad-9f1c-4da9-b52a3a01731c.md)
#### Other resources


 [SelectNamesDialog Object Members](0f5546af-f89a-8a8b-ced9-a2d646bf9634.md)
