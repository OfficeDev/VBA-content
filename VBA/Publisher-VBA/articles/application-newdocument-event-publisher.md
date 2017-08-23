---
title: "Событие Application.NewDocument (издатель)"
keywords: vbapb10.chm268435462
f1_keywords: vbapb10.chm268435462
ms.prod: publisher
api_name: Publisher.Application.NewDocument
ms.assetid: 629cf55c-5134-4207-14df-143b517b9f36
ms.date: 06/08/2017
ms.openlocfilehash: b0643150d7d9bc2c8ab5466e310bb5df2d1cc7d1
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationnewdocument-event-publisher"></a>Событие Application.NewDocument (издатель)

Происходит при создании новой публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **NewDocument** ( **_Doc_**),

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Doc|Обязательное свойство.| **Документ**|Новый документ.|

## <a name="example"></a>Пример

В этом примере выводится сообщение при создании новой публикации.


```vb
Private Sub appPub_NewDocument(ByVal Doc As Document) 
 MsgBox "This is a new publication." 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

