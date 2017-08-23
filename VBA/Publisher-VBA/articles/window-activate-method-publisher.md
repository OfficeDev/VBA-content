---
title: "Метод Window.Activate (издатель)"
keywords: vbapb10.chm262162
f1_keywords: vbapb10.chm262162
ms.prod: publisher
api_name: Publisher.Window.Activate
ms.assetid: 9bd17970-d038-33de-18ad-139bd9fdb8e8
ms.date: 06/08/2017
ms.openlocfilehash: 2ae05af90e90e1c0af876e444be43c76b0b86a49
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="windowactivate-method-publisher"></a>Метод Window.Activate (издатель)

Активирует окно или объекта OLE.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Активация**

 переменная _expression_A, представляющий объект **Window** .


### <a name="return-value"></a>Возвращаемое значение

Значение Nothing


## <a name="remarks"></a>Заметки

Publisher выполняется в одном окне, с помощью метода **активировать** с помощью объекта **Window** , поэтому Publisher активное приложение.


## <a name="example"></a>Пример

В следующем примере создается Publisher активное приложение.


```vb
Application.ActiveWindow.Activate
```

В следующем примере добавляет таблицы Excel в первой страницы публикации, активных и активирует электронной таблицы для редактирования.




```vb
Dim shpSheet As Shape 
 
Set shpSheet = ActiveDocument.Pages(1).Shapes.AddOLEObject _ 
 (Left:=72, Top:=72, ClassName:="Excel.Sheet") 
 
shpSheet.OLEFormat.Activate
```


