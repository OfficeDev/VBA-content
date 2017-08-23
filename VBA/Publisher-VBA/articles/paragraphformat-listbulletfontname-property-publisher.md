---
title: "Свойство ParagraphFormat.ListBulletFontName (издатель)"
keywords: vbapb10.chm5439525
f1_keywords: vbapb10.chm5439525
ms.prod: publisher
api_name: Publisher.ParagraphFormat.ListBulletFontName
ms.assetid: aa0269a1-c5a8-1705-551f-6b1b849701e9
ms.date: 06/08/2017
ms.openlocfilehash: 7ad6ac242d7c6ef342beb9a9e585f7e686bdad66
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatlistbulletfontname-property-publisher"></a>Свойство ParagraphFormat.ListBulletFontName (издатель)

Задает или получает **строку** , представляющую имя шрифта маркера списка из указанного абзацев. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ListBulletFontName**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="remarks"></a>Заметки

Возвращает сообщение «Доступ запрещен», если список не маркированный список.


## <a name="example"></a>Пример

В этом примере проверяется, если тип списка — маркированный список. Если он установлен, **ListBulletFontName** задано значение «Verdana», **ListFontSize** задано значение 24.


```vb
Dim objParaForm As ParagraphFormat 
 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
.TextFrame.TextRange.ParagraphFormat 
 
With objParaForm 
 If .ListType = pbListTypeBullet Then 
 .ListBulletFontName = "Verdana" 
 .ListBulletFontSize = 24 
 End If 
End With 

```


