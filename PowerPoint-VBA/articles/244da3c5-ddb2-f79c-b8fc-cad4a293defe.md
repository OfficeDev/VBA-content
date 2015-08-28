
# View.PrintOut Method (PowerPoint)

 **Last modified:** July 28, 2015

Prints the specified presentation.

## Syntax

 _expression_. **PrintOut**( **_From_**,  **_To_**,  **_PrintToFile_**,  **_Copies_**,  **_Collate_**)

 _expression_A variable that represents a  **View** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|From|Optional| **Long**|The number of the first page to be printed. If this argument is omitted, printing starts at the beginning of the presentation. Specifying the  **To** and **From** arguments sets the contents of the ** [PrintRanges](5c1e9dc1-e30c-bc65-5283-448b95795b11.md)** object and sets the value of the **RangeType** property for the presentation.|
|To|Optional| **Long**|The number of the last page to be printed. If this argument is omitted, printing continues to the end of the presentation. Specifying the  **To** and **From** arguments sets the contents of the ** [PrintRanges](5c1e9dc1-e30c-bc65-5283-448b95795b11.md)** object and sets the value of the **RangeType** property for the presentation.|
|PrintToFile|Optional| **String**|The name of the file to print to. If you specify this argument, the file is printed to a file rather than sent to a printer. If this argument is omitted, the file is sent to a printer.|
|Copies|Optional| **Long**|The number of copies to be printed. If this argument is omitted, only one copy is printed. Specifying this argument sets the value of the  ** [NumberOfCopies](6630ac4d-5c19-ad5f-f557-12e25e198e17.md)** property.|
|Collate|Optional| **MsoTriState**|If this argument is omitted, multiple copies are collated. Specifying this argument sets the value of the  ** [Collate](4cf1d714-6ea2-fce5-340e-202d91ad1137.md)** property.|

## Remarks

The  _Collate_ parameter value can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|Prints all copies of one page before printing the first copy of the next page.|
| **msoTrue**|Prints a complete copy of the presentation before the first page of the next copy is printed.|

## See also


#### Concepts


 [View Object](333e8b59-398d-4575-d37b-bfb1d3503089.md)
#### Other resources


 [View Object Members](3330372c-8497-8cce-981b-3b64700eb915.md)
