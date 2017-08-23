---
title: "Свойство MailMergeDataField.MappedTo (издатель)"
keywords: vbapb10.chm6422566
f1_keywords: vbapb10.chm6422566
ms.prod: publisher
api_name: Publisher.MailMergeDataField.MappedTo
ms.assetid: 067619e8-98fe-d0c2-2f50-96b50cf53de4
ms.date: 06/08/2017
ms.openlocfilehash: e656c967c817d95f03bbde8033789f71c0598c76
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatafieldmappedto-property-publisher"></a>Свойство MailMergeDataField.MappedTo (издатель)

Возвращает имя получателя поля (столбца) в источнике данных главной (список получателей объединенный слияния почты), сопоставленную с **MailMergeDataField** родительский объект. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MappedTo**

 переменная _expression_A, представляет собой объект- **MailMergeDataField** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="remarks"></a>Заметки

Родительский объект **MailMergeDataField** должен представлять поля (столбца) для источника данных, который не является источником основных данных (сочетание всех подключенных источников данных). Свойство **MappedTo** недоступно для полей данных в источнике данных, представленного свойство **DataSource** объекта **слияния** активной объекта **Document** ( `ThisDocument.MailMerge.DataSource`).


