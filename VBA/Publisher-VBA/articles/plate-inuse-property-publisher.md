---
title: "Свойство Plate.InUse (издатель)"
keywords: vbapb10.chm2883602
f1_keywords: vbapb10.chm2883602
ms.prod: publisher
api_name: Publisher.Plate.InUse
ms.assetid: 6c98ada2-ff05-30c9-0043-afbe892dab3d
ms.date: 06/08/2017
ms.openlocfilehash: 65b9b30731f76a245e71663714ad3014a293ba2e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="plateinuse-property-publisher"></a>Свойство Plate.InUse (издатель)

Возвращает **значение True** , если указанный рукописный ввод (представленное форму) используется в публикации. Только для чтения **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Может быть каталогом**

 переменная _expression_A, представляющий объект **формы** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Это свойство соответствует нотацию **Используется** или **Не используется** , перечисленные по каждой рукописного ввода на вкладке **рукописного ввода** диалогового окна **Печать цвета** .


## <a name="example"></a>Пример

В следующем примере циклически просматривает коллекцию форм active публикации, определяет, какие формы представляют краски, которые не используются в публикации и их удаление.


```vb
Sub DeleteUnusedInks() 
 
Dim intCount As Integer 
 
With ActiveDocument.Plates 
 For intCount = .Count To 1 Step -1 
 With .Item(intCount) 
 If .InUse = False Then 
 Debug.Print "Name: " &; .Name 
 .Delete 
 End If 
 End With 
 Next 
End With 
 
End Sub
```


