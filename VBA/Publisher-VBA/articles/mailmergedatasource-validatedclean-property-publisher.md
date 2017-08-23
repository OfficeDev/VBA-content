---
title: "Свойство MailMergeDataSource.ValidatedClean (издатель)"
keywords: vbapb10.chm6291497
f1_keywords: vbapb10.chm6291497
ms.prod: publisher
api_name: Publisher.MailMergeDataSource.ValidatedClean
ms.assetid: 652d2c25-dd15-7431-897b-b17b171b10ea
ms.date: 06/08/2017
ms.openlocfilehash: cc501ad20d987f6fdc1e3d4781ab1f234f9ee382
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasourcevalidatedclean-property-publisher"></a>Свойство MailMergeDataSource.ValidatedClean (издатель)

Указывает, будет ли все адреса получателей в родительский объект **вывода** успешно прошли проверку, и следует ли внесения изменений в список с момента последней проверки, который требуется список проверяемые еще раз. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ValidatedClean**

 переменная _expression_A, представляющий объект **вывода** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

При создании надстройки для Microsoft Publisher, проверяет адреса получателей и обслуживает собственные источники данных,-связи можно задать значение свойства **ValidatedClean** значение **True,** после успешной проверки.

Значение свойства **ValidatedClean** не сохраняется в файле издателя и имеет значение **False** по умолчанию при открытии публикации.

Publisher сбрасывает значение свойства **ValidatedClean** **False** при добавить новый источник данных, изменить параметр filter или изменение Настройка сортировки.


