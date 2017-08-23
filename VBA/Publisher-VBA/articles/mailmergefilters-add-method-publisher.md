---
title: "Метод MailMergeFilters.Add (издатель)"
keywords: vbapb10.chm6750212
f1_keywords: vbapb10.chm6750212
ms.prod: publisher
api_name: Publisher.MailMergeFilters.Add
ms.assetid: ab114dda-d144-7c5f-88b0-930cadcf53db
ms.date: 06/08/2017
ms.openlocfilehash: ddfcf25dee80cdc7f5f4866ebd5cbba519fbaab5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergefiltersadd-method-publisher"></a>Метод MailMergeFilters.Add (издатель)

Добавление нового условия фильтра на указанный объект **MailMergeFilters** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Добавление** ( **_Столбец_**, **_сравнения_**, **_совместно_**, **_bstrCompareTo_**, **_DeferUpdate_**)

 переменная _expression_A, представляет собой объект- **MailMergeFilters** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Столбец|Обязательное свойство.| **String**|Имя таблицы в источнике данных.|
|Comparison|Обязательное свойство.| **MsoFilterComparison**|Способ фильтрации данных в таблице.|
|Совместно|Обязательное свойство.| **MsoFilterConjunction**| Как этот фильтр относится к другим фильтров в объекте **MailMergeFilters** .|
|bstrCompareTo|Необязательный| **String**|Если аргумент **сравнения** , что-то отличного от **msoFilterComparisonIsBlank** или **msoFilterComparisonIsNotBlank**, строка, с которым сравнивается данных в таблице.|
|DeferUpdate|Необязательный| **Boolean**| **Значение true,** для создания очереди фильтры и применить их при вызове метода **ApplyFilter** . **Значение false,** Чтобы немедленно применить условия фильтра. Значение по умолчанию — **False**.|

## <a name="remarks"></a>Заметки

Сравнение может иметь одно из следующих констант **MsoFilterComparison** .



| **msoFilterComparisonContains**|| **msoFilterComparisonEqual**|| **msoFilterComparisonGreaterThan**|| **msoFilterComparisonGreaterThanEqual**|| **msoFilterComparisonIsBlank**|| **msoFilterComparisonIsNotBlank**|| **msoFilterComparisonLessThan**|| **msoFilterComparisonLessThanEqual**|| **msoFilterComparisonNotContains**|| **msoFilterComparisonNotEqual**| Совместно может иметь одно из следующих констант **MsoFilterConjunction** .



| **msoFilterConjunctionAnd**|| **msoFilterConjunctionOr**|

