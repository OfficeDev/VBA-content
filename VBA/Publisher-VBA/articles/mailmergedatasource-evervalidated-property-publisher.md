---
title: "Свойство MailMergeDataSource.EverValidated (издатель)"
keywords: vbapb10.chm6291496
f1_keywords: vbapb10.chm6291496
ms.prod: publisher
api_name: Publisher.MailMergeDataSource.EverValidated
ms.assetid: f87980c8-d327-9313-fa6d-efdfaecb0e35
ms.date: 06/08/2017
ms.openlocfilehash: 6a99fd112d89f0dd6a0f5bdeb1d61ea12eab1702
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasourceevervalidated-property-publisher"></a>Свойство MailMergeDataSource.EverValidated (издатель)

Указывает является ли список адресов получателей в родительский объект **вывода** когда-либо проверки. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **EverValidated**

 переменная _expression_A, представляющий объект **вывода** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

 Начальное значение **EverValidated** имеет **значение False**.

При создании надстройки для Microsoft Publisher, проверяет адреса получателей и обслуживает собственные источники данных,-связи можно задать значение свойства **EverValidated** значение **True,** после успешной проверки.

Значение свойства **EverValidated** сохраняется в файле Microsoft Publisher и доступен в нескольких сеансах Publisher.


