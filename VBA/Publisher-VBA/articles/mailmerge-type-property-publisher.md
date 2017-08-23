---
title: "Свойство MailMerge.Type (издатель)"
keywords: vbapb10.chm6225945
f1_keywords: vbapb10.chm6225945
ms.prod: publisher
api_name: Publisher.MailMerge.Type
ms.assetid: cd31c23f-4059-c6ae-851a-ec9b7f107724
ms.date: 06/08/2017
ms.openlocfilehash: 135db53eccdc8160d2e6d32ebbc73dc3dd589cc1
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergetype-property-publisher"></a>Свойство MailMerge.Type (издатель)

Получает или задает тип слияния почты, представленного объектом **подпапку** родительской. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Тип**

 _expression_An выражение, возвращающее объект **слияния** .


### <a name="return-value"></a>Возвращаемое значение

 **PbMergeType**


## <a name="remarks"></a>Заметки

Возможные значения для свойства **типа** объявлен в перечислении **PbMergeType** и показаны в следующей таблице.



|**Константы**|**Значение**|**Описание**|
|:-----|:-----|:-----|
| **pbCatalogMerge**|3|Объединение в каталог|
| **pbEmailMerge**|4|Слияние почты|
| **pbMailMerge**|2|Слияние почты|
| **pbMergeDefault**|0|По умолчанию объединения|

