---
title: "Объект MailMergeFilters (издатель)"
keywords: vbapb10.chm6815743
f1_keywords: vbapb10.chm6815743
ms.prod: publisher
api_name: Publisher.MailMergeFilters
ms.assetid: 3a91c67f-6cc2-1d67-3382-04ead84f6f09
ms.date: 06/08/2017
ms.openlocfilehash: 71b30815aee9e2e36b2dcb4ba81ff0a91a0f0b4d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergefilters-object-publisher"></a>Объект MailMergeFilters (издатель)

Представляет все фильтры, применяемые к источнику данных, подключенного к слияния почты и публикации слиянием каталога. Объект **MailMergeFilters** состоит из **MailMergeFilterCriterion** объектов.
 


## <a name="example"></a>Пример

Используйте метод **[Add](mailmergefilters-add-method-publisher.md)** объекта **MailMergeFilters** для добавления нового условия фильтра к запросу. В этом примере добавляет новую строку в строке запроса, а затем применяет объединенный фильтра к источнику данных. В этом примере предполагается, что источник данных подключен к активной публикации.
 

 

```
Sub FilterDataSource() 
 With ActiveDocument.MailMerge.DataSource 
 .Filters.Add Column:="Region", _ 
 Comparison:=msoFilterComparisonIsBlank, _ 
 Conjunction:=msoFilterConjunctionAnd 
 .ApplyFilter 
 End With 
End Sub
```

Используйте метод **[элемента](mailmergefilters-item-method-publisher.md)** для доступа к отдельным условиям фильтра. В этом примере циклически просматривает все критерии фильтра и его удаление из слияния почты все записи, которые не равно «WA» изменяется при обнаружении со значением «Область». В этом примере предполагается, что источник данных подключен к активной публикации.
 

 



```
Sub SetQueryCriterion() 
 Dim intItem As Integer 
 With ActiveDocument.MailMerge.DataSource.Filters 
 For intItem = 1 To .Count 
 With .Item(intItem) 
 If .Column = "Region" Then 
 .Comparison = msoFilterComparisonNotEqual 
 .CompareTo = "WA" 
 If .Conjunction = "Or" Then .Conjunction = "And" 
 End If 
 End With 
 Next 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Добавление](mailmergefilters-add-method-publisher.md)|
|[Delete](mailmergefilters-delete-method-publisher.md)|
|[Элемент](mailmergefilters-item-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](mailmergefilters-application-property-publisher.md)|
|[Count](mailmergefilters-count-property-publisher.md)|
|[Создатель](mailmergefilters-creator-property-publisher.md)|
|[Родительский раздел](mailmergefilters-parent-property-publisher.md)|

