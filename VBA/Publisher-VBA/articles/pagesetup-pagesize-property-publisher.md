---
title: "Свойство PageSetup.PageSize (издатель)"
keywords: vbapb10.chm6946850
f1_keywords: vbapb10.chm6946850
ms.prod: publisher
api_name: Publisher.PageSetup.PageSize
ms.assetid: b0605215-5d91-e26e-d3c5-98254cf30044
ms.date: 06/08/2017
ms.openlocfilehash: c72b3fa2f5c2fd2a571d25da337cfb46fd339f88
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagesetuppagesize-property-publisher"></a>Свойство PageSetup.PageSize (издатель)

Получает или задает размер пустая страница для текущей публикации. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PageSize**

 переменная _expression_A, представляет собой объект- **PageSetup** .


### <a name="return-value"></a>Возвращаемое значение

PageSize


## <a name="remarks"></a>Заметки

Размер пустая страница, представленного объектом **PageSize** возвращаются или задаются свойством **PageSize** соответствует одному значков, отображаемых в разделе **Пустая страница размеры** в диалоговом окне **Параметры страницы** в интерфейсе пользователя Microsoft Publisher.


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как задать размер пустая страница для текущей публикации. В примере задается размер пустая страница, чтобы «Индекс карты» — размер пустую страницу по номеру индекса 130 в коллекции **AvailablePageSizes** . В разделе **[AvailablePageSizes](pagesetup-availablepagesizes-property-publisher.md)** свойство пример того, как создать текстовый файл, в котором перечислены все размеры страницы, доступные в текущей публикации и их соответствующих значений индекса.


```vb
Public Sub PageSize_Example() 
 
 ThisDocument.PageSetup.PageSize = ThisDocument.PageSetup.AvailablePageSizes.Item(130) 
 
End Sub
```


