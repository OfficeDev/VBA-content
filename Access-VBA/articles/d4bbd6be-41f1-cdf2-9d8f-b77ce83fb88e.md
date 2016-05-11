
# StreamOpenOptionsEnum

 **Last modified:** March 09, 2015

 _ **Applies to:** Access 2013 | Access 2016_



Specifies options for opening a [Stream](d49b1514-e0b4-0aca-d5c2-8266f3f4fe65.md) object. The values can be combined with an OR operation.


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**adOpenStreamAsync**|1|Opens the  **Stream** object in asynchronous mode.|
|**adOpenStreamFromRecord**|4|Identifies the contents of the  _Source_ parameter to be an already open[Record](817aaf13-78d4-1134-aa94-997e92077c22.md) object. The default behavior is to treat _Source_ as a URL that points directly to a node in a tree structure. The default stream associated with that node is opened.|
|**adOpenStreamUnspecified**|-1|Default. Specifies opening the  **Stream** object with default options.|
 **ADO/WFC Equivalent**
These constants do not have ADO/WFC equivalents.
