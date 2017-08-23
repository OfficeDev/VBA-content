---
title: "Свойство EmailMergeEnvelope.Bcc (издатель)"
keywords: vbapb10.chm9043974
f1_keywords: vbapb10.chm9043974
ms.prod: publisher
api_name: Publisher.EmailMergeEnvelope.Bcc
ms.assetid: 1d846fac-d93c-6a20-ce3b-090525dbbfe1
ms.date: 06/08/2017
ms.openlocfilehash: be48f830bdf9e887257076090766d232653e81d4
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="emailmergeenvelopebcc-property-publisher"></a>Свойство EmailMergeEnvelope.Bcc (издатель)

Получает или задает список адресов электронной почты, получающих скрытую копию (BCC) сообщения электронной почты, разделенных точкой с запятой. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Скрытой копии**

 переменная _expression_A, представляющий объект **EmailMergeEnvelope** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="remarks"></a>Заметки

Установите свойство **Bcc** в строке адреса электронной почты, разделенных точкой с запятой, как показано в следующем примере.


```vb
 MailMerge.EmailMergeEnvelope.Bcc = "name1@address1;name2@address2;name3@address3;..."
```


