---
title: "Свойство EmailMergeEnvelope.Priority (издатель)"
keywords: vbapb10.chm9043976
f1_keywords: vbapb10.chm9043976
ms.prod: publisher
api_name: Publisher.EmailMergeEnvelope.Priority
ms.assetid: 21c4c33f-d211-7ca5-364b-be9ad4d3f187
ms.date: 06/08/2017
ms.openlocfilehash: fa13bd7271bb2347a0c5b0ad6f99149c7dfe7e00
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="emailmergeenvelopepriority-property-publisher"></a>Свойство EmailMergeEnvelope.Priority (издатель)

Получает или задает приоритет сообщения электронной почты объединенные, представленного объектом **EmailMergeEnvelope** родительского. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Приоритет**

 переменная _expression_A, представляющий объект **EmailMergeEnvelope** .


### <a name="return-value"></a>Возвращаемое значение

pbEmailMergePriority


## <a name="remarks"></a>Заметки

Возможные значения для свойства **Priority** объявления в перечислении **pbEmailMergePriority** и показаны в следующей таблице.



|**Константы**|**Значение**|**Описание**|
|:-----|:-----|:-----|
| **pbPriorityNone**|0|Не установлен приоритет|
| **pbPriorityLow**|2|Низкая важность|
| **pbPriorityHigh**|1|Высокий приоритет|

