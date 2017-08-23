---
title: "Свойство MailMerge.ViewMailMergeFieldCodes (издатель)"
keywords: vbapb10.chm6225928
f1_keywords: vbapb10.chm6225928
ms.prod: publisher
api_name: Publisher.MailMerge.ViewMailMergeFieldCodes
ms.assetid: 05b5e6e2-10ae-c6e0-3214-7016295703e2
ms.date: 06/08/2017
ms.openlocfilehash: 7f67265193cb1d6d3d22c18aa24c3c0786fc6294
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergeviewmailmergefieldcodes-property-publisher"></a>Свойство MailMerge.ViewMailMergeFieldCodes (издатель)

 **Значение true,** Если имена полей слияния отображаются в публикации слияния почты; **Значение false,** Если отображаются данные из текущей записи. Чтение и запись **типа Boolean**. .


## <a name="syntax"></a>Синтаксис

 _выражение_. **ViewMailMergeFieldCodes**

 переменная _expression_A, представляет собой объект- **слияния** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Если активная публикация не публикации слияния, с помощью этого свойства не оказывает влияния.


## <a name="example"></a>Пример

В этом примере скрывает коды полей слияния почты в активной публикации.


```vb
ActiveDocument.MailMerge.ViewMailMergeFieldCodes = False 

```


