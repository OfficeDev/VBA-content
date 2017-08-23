---
title: "Свойство WebCommandButton.DataRetrievalMethod (издатель)"
keywords: vbapb10.chm3932166
f1_keywords: vbapb10.chm3932166
ms.prod: publisher
api_name: Publisher.WebCommandButton.DataRetrievalMethod
ms.assetid: 81b89a3b-dcc5-c2b5-fbc4-6e02b587bc42
ms.date: 06/08/2017
ms.openlocfilehash: 8eaeaf12ac4af1d9cc7b162862dd4ad5ca1b76f7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webcommandbuttondataretrievalmethod-property-publisher"></a>Свойство WebCommandButton.DataRetrievalMethod (издатель)

Задает или возвращает обработки **PbSubmitDataRetrievalMethodType** , представляющий способ данные из веб-форму. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **DataRetrievalMethod**

 переменная _expression_A, представляет собой объект- **WebCommandButton** .


### <a name="return-value"></a>Возвращаемое значение

PbSubmitDataRetrievalMethodType


## <a name="remarks"></a>Заметки

Значение свойства **DataRetrievalMethod** может иметь одно из **[PbSubmitDataRetrievalMethodType](pbsubmitdataretrievalmethodtype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В этом примере Microsoft Publisher для обработки данных на веб-форму в текущей публикации, отправив сообщение электронной почты на указанный адрес электронной почты.


```vb
Sub WebFormData() 
 With ThisDocument.Pages(1).Shapes(1).WebCommandButton 
 .DataRetrievalMethod = pbSubmitDataRetrievalEmail 
 .EmailAddress = "someone@example.com" 
 .EmailSubject = "Web form data" 
 End With 
End Sub
```


