
# Characters.CharPropsRow Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Returns the index of the row in the Character section of a ShapeSheet window that contains character formatting information for a  **Characters** object. Read-only.


## Syntax

 _expression_. **CharPropsRow**( **_BiasLorR_**)

 _expression_An expression that returns a  **Characters** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|BiasLorR|Required| **Integer**|The direction of the search.|

### Return Value

Integer


## Remarks

If the formatting of the  **Characters** object is represented by more than one row in the Character section of the ShapeSheet window, (in other words, if the **Characters** object consists of a section of text that spans characters that are formatted differently), the **CharPropsRow** property returns -1. Under these circumstances, Microsoft Visio ignores the value of theBiasLorR argument. (Characters that have the same character formatting share the same row in the ShapeSheet. Visio creates a new ShapeSheet row only when character formatting changes, for example from bold to italic.)

If the  **Characters** object spans several characters within the same character property row, **CharPropsRow** returns the index of that row. In this case as well, Visio ignores theBiasLorR argument.

If the  **Characters** object represents an insertion point rather than a sequence of characters (that is, if its **Begin** and **End** properties return the same value), use theBiasLorR argument to determine which row index to return.



|**Constant **|**Value **|
|:-----|:-----|
| **visBiasLetVisioChoose**|0 |
| **visBiasLeft**|1 |
| **visBiasRight**|2 |
Specify  **visBiasLeft** for the row that covers character formatting for the character to the left of the insertion point, or **visBiasRight** for the row that covers character formatting for the character to the right of the insertion point.

If you specify  **visBiasLetVisioChoose**, Visio uses the same logic it would apply to new text typed in the user interface starting at the insertion point. Usually, that means that Visio will apply the character formatting of the character to the left of the insertion point to the new text, so  **CharPropsRow** will return the same value it would if passed **visBiasLeft**. (For an explanation of the meaning of "left" in this context, see the following note.) However, if the insertion point is at the beginning of a new paragraph,  **CharPropsRow** returns the value it would return if passed **visBiasRight**. 


 **Note**  In the context of a  **Characters** object, "left" means logically prior. In other words, one character is to the "left" of another if it would have been typed first in the course of normal writing. It is necessary to make this distinction because in some languages, characters are normally written from right to left, and not from left to right.

