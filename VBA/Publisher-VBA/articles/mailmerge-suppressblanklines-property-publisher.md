---
title: "Свойство MailMerge.SuppressBlankLines (издатель)"
keywords: vbapb10.chm6225927
f1_keywords: vbapb10.chm6225927
ms.prod: publisher
api_name: Publisher.MailMerge.SuppressBlankLines
ms.assetid: 3b41e0c0-8588-e86a-77ed-90c4692c03dc
ms.date: 06/08/2017
ms.openlocfilehash: f94ed1107a6bb92f1ac9736aa24ae1e8d1e1afdc
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergesuppressblanklines-property-publisher"></a>Свойство MailMerge.SuppressBlankLines (издатель)

 **Значение true,** чтобы не отображать пустые строки при пустые поля слияния в основной документ. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SuppressBlankLines**

 переменная _expression_A, представляет собой объект- **слияния** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В этом примере блокирует отображение пустых строк в активной публикации при полей слияния почты являются пустыми. В этом примере предполагается, что источник данных слияния почты подключенный к активной публикации.


```vb
Sub SuppressBlankLines() 
 ActiveDocument.MailMerge.SuppressBlankLines = True 
End Sub
```


