---
title: "Свойство Field.Code (издатель)"
keywords: vbapb10.chm6094851
f1_keywords: vbapb10.chm6094851
ms.prod: publisher
api_name: Publisher.Field.Code
ms.assetid: bb2f3b23-dea1-bdfb-90bf-4b4ea09570f6
ms.date: 06/08/2017
ms.openlocfilehash: 1ff4c138ccda3e305550416df5408d6fdd163596
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fieldcode-property-publisher"></a>Свойство Field.Code (издатель)

Возвращает **строку** , представляющую текст, отображаемый при Просмотр страницы имеет значение для отображения кодов полей. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Код**

 переменная _expression_A, представляющий объект **поля** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В этом примере циклически просматривает все поля в активной публикации и затем отображает сообщение как, чтобы ли строка «www» найден в коде поля.


```vb
Sub FindWWWHyperlinks() 
 Dim intItem As Integer 
 Dim intField As Integer 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange.Fields 
 Do 
 intItem = intItem + 1 
 If InStr(1, .Item(intItem).Code, "www") > 0 Then 
 intField = intField + 1 
 End If 
 Loop Until intItem = .Count 
 End With 
 
 If intField > 0 Then 
 MsgBox "You have " &; intField &; " World Wide Web " &; _ 
 "hyperlinks in your publication." 
 Else 
 MsgBox "You have no hyperlink fields in your publication." 
 End If 
End Sub
```


