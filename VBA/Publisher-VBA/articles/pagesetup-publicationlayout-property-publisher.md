---
title: "Свойство PageSetup.PublicationLayout (издатель)"
keywords: vbapb10.chm6946839
f1_keywords: vbapb10.chm6946839
ms.prod: publisher
ms.assetid: 6c476789-577d-2088-37dc-bcaed25cd219
ms.date: 06/08/2017
ms.openlocfilehash: e7153d8ec76ec6f4257b69f74c7a07dd5613b94a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagesetuppublicationlayout-property-publisher"></a>Свойство PageSetup.PublicationLayout (издатель)

Возвращает или задает константа [Перечисления PbPublicationLayout (издатель)](pbpublicationlayout-enumeration-publisher.md) , который указывает расположение публикации. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PublicationLayout**

 переменная _expression_A, представляет собой объект- **PageSetup** .


## <a name="return-value"></a>Возвращаемое значение

 **PBPUBLICATIONLAYOUT**


## <a name="remarks"></a>Заметки

С помощью свойства **PublicationLayout** задать макет публикации эквивалентно параметру макет из списка в диалоговом окне **Параметры страницы** .


## <a name="example"></a>Пример

В следующем примере задается макет active публикации для **pbLayoutBusinessCardUS**, который по умолчанию определяет ширину страницы 3,5 дюйма и высота страницы 2 дюйма.


```vb
With ActiveDocument.PageSetup
    .PublicationLayout = pbLayoutBusinessCardUS
End With

```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект PageSetup (издатель)](pagesetup-object-publisher.md)

