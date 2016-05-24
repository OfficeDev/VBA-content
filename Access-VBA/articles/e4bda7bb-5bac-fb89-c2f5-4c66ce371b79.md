
# Cannot delete this index or table. It is either the current index or is used in a relationship. (Error 3281)

 **Last modified:** December 30, 2015

 _ **Applies to:** Access 2013 | Access 2016_

If the index or table is used in a relationship, you must delete the relationship before you can delete the index or table. If the index is specified as the current index by the  **Index** property, you must set the property to a different index before you can delete the index.

