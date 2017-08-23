---
title: "Свойство ParagraphFormat.ListBulletText (издатель)"
keywords: vbapb10.chm5439523
f1_keywords: vbapb10.chm5439523
ms.prod: publisher
api_name: Publisher.ParagraphFormat.ListBulletText
ms.assetid: fa80957a-be91-398f-a24f-5a0449a9466f
ms.date: 06/08/2017
ms.openlocfilehash: 3163b35ce46bce49710f1f91294d6139319a2fc9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatlistbullettext-property-publisher"></a>Свойство ParagraphFormat.ListBulletText (издатель)

Возвращает **строку** , представляющую текст маркированный список из указанного абзацев. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ListBulletText**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="remarks"></a>Заметки

Свойство **ListBulletText** только для одного символа.

Это свойство доступно только для чтения. Чтобы задать свойство **ListBulletText** маркированного списка, используйте метод **SetListType** .

Возвращает сообщение «Доступ запрещен», если список не маркированный список.


## <a name="example"></a>Пример

В этом примере проверяется, если тип списка — маркированный список. Если он установлен, тест выполняется, что текст маркированный список задано значение «*». Если он не установлен, метод **SetListType** вызван и передается как параметр pbListType **pbListTypeBullet** и "*" в качестве BulletText параметра.


```vb
Dim objParaForm As ParagraphFormat 
 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
.TextFrame.TextRange.ParagraphFormat 
 
With objParaForm 
 If .ListType = pbListTypeBullet Then 
 If Not .ListBulletText = "*" Then 
 .SetListType pbListTypeBullet, "*" 
 End If 
 End If 
End With 
 

```


