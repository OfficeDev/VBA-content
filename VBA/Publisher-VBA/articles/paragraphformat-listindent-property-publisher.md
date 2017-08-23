---
title: "Свойство ParagraphFormat.ListIndent (издатель)"
keywords: vbapb10.chm5439522
f1_keywords: vbapb10.chm5439522
ms.prod: publisher
api_name: Publisher.ParagraphFormat.ListIndent
ms.assetid: b42000ea-0636-88cf-b7ed-c71384a2b0d5
ms.date: 06/08/2017
ms.openlocfilehash: 284c3661431d9eb7a27ba68eef82efe9d8376f3c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatlistindent-property-publisher"></a>Свойство ParagraphFormat.ListIndent (издатель)

Возвращает или задает **один** , который представляет значение отступ списка (в пунктах) для указанного объекта **ParagraphFormat** . Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ListIndent**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="example"></a>Пример

В этом примере задается свойство **ListIndent** объекта **ParagraphFormat** 0,25 дюйма. Метод **InchesToPoints не была назначена** используется для преобразования дюймов в пунктах.


```vb
Dim objParaForm As ParagraphFormat 
 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
.TextFrame.TextRange.ParagraphFormat 
 
With objParaForm 
 .ListIndent = InchesToPoints(0.25) 
End With 

```


