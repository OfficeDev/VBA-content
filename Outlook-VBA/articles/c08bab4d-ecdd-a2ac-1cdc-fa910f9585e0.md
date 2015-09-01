
# OlViewSaveOption Enumeration (Outlook)

 **Last modified:** July 28, 2015

Specifies the folders in which the view is available and the read permissions attached to the view.


|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olViewSaveOptionAllFoldersOfType**|2|Indicates that the view is available in all folders of the same type.|
| **olViewSaveOptionThisFolderEveryone**|0|Indicates that the view is only available in the current folder and is available to all users.|
| **olViewSaveOptionThisFolderOnlyMe**|1|Indicates that the view is only available in the current folder and is only available to the current Outlook user.|

## Remarks

Used by the  **Copy** method and **SaveOption** property of **View** objects.

