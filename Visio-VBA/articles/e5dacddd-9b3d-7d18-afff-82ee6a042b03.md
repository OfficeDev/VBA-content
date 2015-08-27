
# ServerPublishOptions.GetPagesToPublish Method (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Returns an array of pages that are set to be published to a server.


## Syntax

 _expression_. **GetPagesToPublish**( **_Flags_**, **_PublishPages_**,  **_NamesArray()_**)

 _expression_A variable that represents a  ** [ServerPublishOptions](69e71212-4ca3-9fa6-6af3-8f07af540140.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Flags|Required| ** [VisLangFlags](9654b6db-072a-6bcb-929d-05d18cb96009.md)**|Out parameter. Indicates whether universal or local page names are returned in NamesArray. See Remarks for possible values.|
|PublishPages|Required| ** [VisPublishPages](a638bea0-67e5-0fd1-1984-ffafb37afcb2.md)**|Out parameter. Indicates whether all pages or selected pages are set to be published. See Remarks for possible values.|
|NamesArray()|Required| **String**|Out parameter. Returns the names of the pages set to be published.|

### Return Value

 **Nothing**


## Remarks

The  _Flags_ parameter can be one of the following **VisLangFlags** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visLangLocal**|0|The page name is a local name.|
| **visLangUniversal**|1|The page name is a universal name.|
The  _PublishPages_ parameter can be one of the following **VisPublishPages** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visPublishPageAll**|0|Publish all pages.|
| **visPublishPageSelect**|1|Publish selected pages.|
