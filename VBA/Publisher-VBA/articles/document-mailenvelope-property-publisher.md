---
title: "Свойство Document.MailEnvelope (издатель)"
keywords: vbapb10.chm196627
f1_keywords: vbapb10.chm196627
ms.prod: publisher
api_name: Publisher.Document.MailEnvelope
ms.assetid: 3c4c734a-6725-5f6e-ed0a-5b19e4e642bd
ms.date: 06/08/2017
ms.openlocfilehash: 53c7aaf74cef4491722aeded8ff3f1c31b2f8387
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentmailenvelope-property-publisher"></a>Свойство Document.MailEnvelope (издатель)

Возвращает объект, представляющий заголовке сообщения электронной почты для публикации на **MsoEnvelope** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **MailEnvelope**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

MsoEnvelope


## <a name="remarks"></a>Заметки

Свойство **MailEnvelope** доступен только в случае, если свойство **[EnvelopeVisible](document-envelopevisible-property-publisher.md)** задано значение **True**.


## <a name="example"></a>Пример

В этом примере задает комментариев в заголовке электронной почты active публикации. В этом примере предполагается, что **EnvelopeVisible** свойство значение **True**.


```vb
Sub HeaderComments() 
 ActiveDocument.MailEnvelope.Introduction = _ 
 "Please review this publication and let me know " &; _ 
 "what you think. I need your input by Friday." &; _ 
 " Thanks." 
End Sub
```


