
# ForeColor Property (Microsoft Forms)

 **Last modified:** July 28, 2015


Specifies the  [foreground color](7ce2c60f-29fb-96e2-2516-73c99a6e7cff.md) of an object.
 **Syntax**
 _object_. **ForeColor** [= _Long_]
The  **ForeColor** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Long_|Optional. A value or constant that determines the foreground color of an object.|
 **Settings**
You can use any integer that represents a valid color. You can also specify a color by using the  [RGB](7ce2c60f-29fb-96e2-2516-73c99a6e7cff.md) function with red, green, and blue color components. The value of each color component is an integer that ranges from zero to 255. For example, you can specify teal blue as the integer value 4966415 or as red, green, and blue color components 15, 200, 75.
 **Remarks**
Use the  **ForeColor** property for controls on forms to make them easy to read or to convey a special meaning. For example, if a text box reports the number of units in stock, you can change the color of the text when the value falls below the reorder level.
For a  **ScrollBar** or **SpinButton**,  **ForeColor** sets the color of the arrows. For a **Frame**,  **ForeColor** changes the color of the caption. For a **Font** object, **ForeColor** determines the color of the text.
