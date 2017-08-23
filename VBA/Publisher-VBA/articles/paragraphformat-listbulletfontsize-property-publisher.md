---
title: "Свойство ParagraphFormat.ListBulletFontSize (издатель)"
keywords: vbapb10.chm5439524
f1_keywords: vbapb10.chm5439524
ms.prod: publisher
api_name: Publisher.ParagraphFormat.ListBulletFontSize
ms.assetid: 1ff1de0f-afcc-cc9c-bf45-d745695db89b
ms.date: 06/08/2017
ms.openlocfilehash: ffbf32f09b79489d41648cc68f82943299b442a7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatlistbulletfontsize-property-publisher"></a>Свойство ParagraphFormat.ListBulletFontSize (издатель)

Задает или получает **единого** , представляющее размер шрифта маркера списка из указанного абзацев. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ListBulletFontSize**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Возвращает сообщение «Доступ запрещен», если список не маркированный список.


## <a name="example"></a>Пример

В этом примере проверяется, если тип списка — маркированный список. Если он установлен, **ListFontSize** задано значение 24, **ListBulletFontName** задано значение «Verdana».


```vb
Dim objParaForm As ParagraphFormat 
 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
.TextFrame.TextRange.ParagraphFormat 
 
With objParaForm 
 If .ListType = pbListTypeBullet Then 
 .ListBulletFontSize = 24 
 .ListBulletFontName = "Verdana" 
 End If 
End With 
 
 

```


