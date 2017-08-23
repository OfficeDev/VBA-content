---
title: "Свойство PageSetup.PageHeight (издатель)"
keywords: vbapb10.chm6946821
f1_keywords: vbapb10.chm6946821
ms.prod: publisher
api_name: Publisher.PageSetup.PageHeight
ms.assetid: 1ef153e2-5d13-d896-cd69-2066efa2f8ef
ms.date: 06/08/2017
ms.openlocfilehash: 71cbdb855792007f8fdb081a94b6e4d138b7f62a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagesetuppageheight-property-publisher"></a>Свойство PageSetup.PageHeight (издатель)

Возвращает или задает **Variant** , который представляет высоту страниц в публикации. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PageHeight**

 переменная _expression_A, представляет собой объект- **PageSetup** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Числовые значения вычисляются как точки. Строковые значения можно в любое устройство, поддерживаемый Microsoft Publisher (например, «2,5 дюйма»). Допустимый диапазон допустимых значений — от 0 до различие между высота листа и высота страницы.


## <a name="example"></a>Пример

В этом примере указывается высота пять дюйма для страниц в активной публикации.


```vb
Public Sub PageHeight_Example() 
 ActiveDocument.PageSetup.PageHeight = InchesToPoints(5) 
End Sub
```


