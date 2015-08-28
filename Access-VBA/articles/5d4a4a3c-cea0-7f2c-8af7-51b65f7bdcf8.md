
# DoCmd.RunCommand Method (Access)

 **Last modified:** July 28, 2015

The  **RunCommand** method runs a built-in command.

## Syntax

 _expression_. **RunCommand**( **_Command_**)

 _expression_A variable that represents a  **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Command|Required| **AcCommand**|An  ** [AcCommand](a78f91cc-3b40-5f45-c737-4d3abb2e979f.md)** constant that specifies the commend to run.|

## Remarks

Each menu and toolbar command in Microsoft Access has an associated constant that you can use with the  **RunCommand** method to run that command from Visual Basic.

You can't use the  **RunCommand** method to run a command on a custom menu or toolbar. You can only use it with built-in menus and toolbars.

The  **RunCommand** method replaces the **DoMenuItem**method of the  **DoCmd** object.


## See also


#### Concepts


 [DoCmd Object](3ce44cca-9979-0a1e-9787-079a52ce528f.md)
#### Other resources


 [DoCmd Object Members](3e7ade9e-86e4-0751-188b-5d31c9101651.md)
