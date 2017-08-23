---
title: "Свойство WebCheckBox.Selected (издатель)"
keywords: vbapb10.chm4325380
f1_keywords: vbapb10.chm4325380
ms.prod: publisher
api_name: Publisher.WebCheckBox.Selected
ms.assetid: ad34871d-474d-70ad-6245-ee5a017839c1
ms.date: 06/08/2017
ms.openlocfilehash: 6f5f327d87a8e4d1cdceb393755140a77a205679
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webcheckboxselected-property-publisher"></a>Свойство WebCheckBox.Selected (издатель)

Указывает, установлен ли флажок Web или переключателя. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Выбранные**

 переменная _expression_A, представляет собой объект- **WebCheckBox** .


## <a name="remarks"></a>Заметки

Значение свойства **Selected** может иметь одно из ** [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


## <a name="example"></a>Пример

В этом примере добавляется новый Web флажок первой страницы публикации, активных и выбирает его.


```vb
Sub AddNewWebCheckBox() 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlCheckBox, Left:=100, _ 
 Top:=100, Width:=100, Height:=12) 
 .WebCheckBox.Selected = msoTrue 
 End With 
End Sub
```


