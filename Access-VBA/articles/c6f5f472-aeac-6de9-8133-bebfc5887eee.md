
# Application.DDERequest Method (Access)

 **Last modified:** July 28, 2015

You can use the  **DDERequest** function over an open dynamic data exchange (DDE) channel to request an item of information from a DDE server application.

## Syntax

 _expression_. **DDERequest**( **_ChanNum_**,  **_Item_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ChanNum|Required| **Variant**|A channel number, the integer returned by the  **DDEInitiate**function.|
|Item|Required| **String**|A string expression that's the name of a data item recognized by the application specified by the  **DDEInitiate** function. Check the application's documentation for a list of possible items.|

### Return Value

String


## Remarks

For example, if you have an open DDE channel between Microsoft Access and Microsoft Excel, you can use the  **DDERequest** function to transfer text from a Microsoft Excel spreadsheet to a Microsoft Access database.

The channum argument specifies the channel number of the desired DDE conversation, and theitem argument identifies which data should be retrieved from the server application. The value of theitem argument depends on the application and topic specified when the channel indicated by thechannum argument is opened. For example, theitem argument may be a range of cells in a Microsoft Excel spreadsheet.

The  **DDERequest** function returns a **Variant** as a string containing the requested information if the request was successful.

The data is requested in alphanumeric text format. Graphics or text in any other format can't be transferred.

If the channum argument isn't an integer corresponding to an open channel, or if the data requested can't be transferred, a run-time error occurs.

If you need to manipulate another application's objects from Microsoft Access, you may want to consider using Automation .


## See also


#### Concepts


 [Application Object](aefb0713-97e6-e2c7-e530-8fd2e1316a55.md)
#### Other resources


 [Application Object Members](3ab5276c-d52a-72a9-244c-ec92ead48811.md)
