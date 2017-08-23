---
title: "Свойство ShapeRange.AlternativeText (издатель)"
keywords: vbapb10.chm2293856
f1_keywords: vbapb10.chm2293856
ms.prod: publisher
api_name: Publisher.ShapeRange.AlternativeText
ms.assetid: 94cbb99b-3b35-76bb-e269-db8295b84f2f
ms.date: 06/08/2017
ms.openlocfilehash: 7f1fde8afc5f1a967780811f75bd2635460af0b3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangealternativetext-property-publisher"></a>Свойство ShapeRange.AlternativeText (издатель)

Возвращает или задает **строку** , представляющую текст, отображаемый в веб-браузере вместо объекта **Shape** , при загрузке объекта **Shape** или после отключения графики. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AlternativeText**

 переменная _expression_A, представляющий объект **ShapeRange** .


## <a name="remarks"></a>Заметки

Максимальная длина свойство **AlternativeText** — 254 символов. Microsoft Publisher возвращает ошибку, если длина текста превышает этот номер.


## <a name="example"></a>Пример

В этом примере задается замещающий текст для выбранной фигуры в активный документ. В этом примере предполагается, что у вас есть публикации, что выбранные фигуры является изображение признакам.


```vb
Public Sub Alternative_Text() 
 
 ' The picture of a duck must be selected. 
 Publisher.ActiveDocument.Selection.ShapeRange _ 
 .AlternativeText = "This is a mallard duck." 
 
End Sub
```


