---
title: "Объект MailMergeFilterCriterion (издатель)"
keywords: vbapb10.chm6881279
f1_keywords: vbapb10.chm6881279
ms.prod: publisher
api_name: Publisher.MailMergeFilterCriterion
ms.assetid: 2814890f-009b-b277-3ea4-c1f167a5e1c9
ms.date: 06/08/2017
ms.openlocfilehash: fe09493cbd5757bb2979b8679c61acdf3290bdb0
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergefiltercriterion-object-publisher"></a>Объект MailMergeFilterCriterion (издатель)

Представляет фильтр должен применяться к вложенные слияния почты и каталогов объединения источника данных. Объект **MailMergeFilterCriterion** , является участником объекта **MailMergeFilters** .
 


## <a name="example"></a>Пример

Каждый фильтр — это строка в строке запроса. Использование свойств **[столбца](mailmergefiltercriterion-column-property-publisher.md)**, **[сравнения](mailmergefiltercriterion-comparison-property-publisher.md)**, **[CompareTo](mailmergefiltercriterion-compareto-property-publisher.md)**и **[совместно](mailmergefiltercriterion-conjunction-property-publisher.md)** для возвращения или задания условия запроса источника данных. В следующем примере изменяется существующий фильтр для удаления из слияния почты все записи, у которых нет поля региона, равное «WA». В этом примере предполагается, что источник данных подключен к активной публикации.
 

 

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


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](mailmergefiltercriterion-application-property-publisher.md)|
|[Столбец](mailmergefiltercriterion-column-property-publisher.md)|
|[CompareTo](mailmergefiltercriterion-compareto-property-publisher.md)|
|[Сравнение](mailmergefiltercriterion-comparison-property-publisher.md)|
|[Совместно](mailmergefiltercriterion-conjunction-property-publisher.md)|
|[Создатель](mailmergefiltercriterion-creator-property-publisher.md)|
|[Index](mailmergefiltercriterion-index-property-publisher.md)|
|[Родительский раздел](mailmergefiltercriterion-parent-property-publisher.md)|

