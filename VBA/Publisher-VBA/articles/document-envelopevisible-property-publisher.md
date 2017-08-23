---
title: "Свойство Document.EnvelopeVisible (издатель)"
keywords: vbapb10.chm196618
f1_keywords: vbapb10.chm196618
ms.prod: publisher
api_name: Publisher.Document.EnvelopeVisible
ms.assetid: 65423c1f-e61b-3c83-4bff-ddd278d97238
ms.date: 06/08/2017
ms.openlocfilehash: 5c9fdb54f320d41c4b236bbde60af9798c3d7a40
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentenvelopevisible-property-publisher"></a>Свойство Document.EnvelopeVisible (издатель)

Возвращает или задает значение **Boolean** , указывающее, отображается ли заголовок сообщения электронной почты в окне публикации. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **EnvelopeVisible**

 переменная _expression_A, представляющий объект **документа** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В этом примере отображаются заголовок сообщения электронной почты для активной публикации.


```vb
ActiveDocument.EnvelopeVisible = True
```


