---
title: "Свойство Font.Size (издатель)"
keywords: vbapb10.chm5373957
f1_keywords: vbapb10.chm5373957
ms.prod: publisher
api_name: Publisher.Font.Size
ms.assetid: 485f68fe-c6d7-8288-042e-fc4c35c37b2d
ms.date: 06/08/2017
ms.openlocfilehash: 3800ee5637730cc32eacc5ea9706cd024ce99084
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontsize-property-publisher"></a>Свойство Font.Size (издатель)

Представляет размер символов в диапазон текста в пунктах. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Размер**

 _expression_An выражение, возвращающее объект **Font** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="example"></a>Пример

В этом примере вставляется в текст и затем задает размер шрифта седьмой слова вставленного текста на 20 точек.


```vb
Sub IncreaseFontSizeOfSelection() 
 With Selection.TextRange 
 .InsertBefore vbLf &; "This is a demonstration of font size." 
 .Words(7).Font.Size = 20 
 End With 
End Sub
```


