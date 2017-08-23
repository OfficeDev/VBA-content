---
title: "Свойство MailMerge.DocumentUpdating (издатель)"
keywords: vbapb10.chm6225925
f1_keywords: vbapb10.chm6225925
ms.prod: publisher
api_name: Publisher.MailMerge.DocumentUpdating
ms.assetid: c65ca4a0-e5eb-d97e-9126-4af86f4e805f
ms.date: 06/08/2017
ms.openlocfilehash: 266f99f38d123274b6b51507d61d2fa4d5673fd7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedocumentupdating-property-publisher"></a>Свойство MailMerge.DocumentUpdating (издатель)

Возвращает или задает значение **Boolean** , указывающее, выполняется ли обновление экрана при выполнении кода слияния почты. Значение по умолчанию — **True** (экрана обновляется). Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **DocumentUpdating**

 переменная _expression_A, представляет собой объект- **слияния** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Отключение обновления документов во время выполнения может ускорить выполнение кода Microsoft Visual Basic. Тем не менее рекомендуется предоставить некоторые указания состояния, пользователя принять во внимание, что программа работает правильно.


## <a name="example"></a>Пример

Следующий пример отключает обновление документа в начале подпрограммы слияния почты и включает его обратно в конце подпрограммы.


```vb
Sub MailMergeProcedure() 
 ActiveDocument.MailMerge.DocumentUpdating = False ' Mail merge code. 
ActiveDocument.MailMerge.DocumentUpdating = True 
End Sub
```


