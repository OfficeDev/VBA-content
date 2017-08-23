---
title: "Метод WebListBoxItems.Selected (издатель)"
keywords: vbapb10.chm4128775
f1_keywords: vbapb10.chm4128775
ms.prod: publisher
api_name: Publisher.WebListBoxItems.Selected
ms.assetid: 2db3b8cb-2922-1cca-9613-67402772ee27
ms.date: 06/08/2017
ms.openlocfilehash: 8ab9d0ee7105a4150e5ef3afc9a7955d3754fba3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="weblistboxitemsselected-method-publisher"></a>Метод WebListBoxItems.Selected (издатель)

Выбирает или отменяет выбор элемента в элемент управления списка Web.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Выбранные** ( **_Индекса_**, **_SelectState_**)

 переменная _expression_A, представляет собой объект- **WebListBoxItems** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Индекс|Обязательное свойство.| **Длинный**|Номер элемента списка Web.|
|SelectState|Обязательное свойство.| **Boolean**| **Значение true,** чтобы выбрать элемент списка.|

## <a name="example"></a>Пример

В этом примере выполняется проверка, что существующий элемент управления полем списка Web позволяет выбирать несколько записей, а затем выбирает два элемента в списке.


```vb
Sub SelectListBoxItem() 
 With ActiveDocument.Pages(1).Shapes(1) _ 
 .WebListBox 
 If .MultiSelect = msoTrue Then 
 With .ListBoxItems 
 .Selected Index:=1, SelectState:=True 
 .Selected Index:=3, SelectState:=True 
 End With 
 End If 
 End With 
End Sub
```


