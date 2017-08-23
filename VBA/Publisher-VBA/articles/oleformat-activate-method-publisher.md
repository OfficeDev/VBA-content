---
title: "Метод OLEFormat.Activate (издатель)"
keywords: vbapb10.chm4456454
f1_keywords: vbapb10.chm4456454
ms.prod: publisher
api_name: Publisher.OLEFormat.Activate
ms.assetid: 43c01633-f624-c5ef-ba2c-d1ff62e91ec5
ms.date: 06/08/2017
ms.openlocfilehash: 36e5bea1e6fa8840733dcf8d7057416c87b64340
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="oleformatactivate-method-publisher"></a>Метод OLEFormat.Activate (издатель)

Активирует окно или объекта OLE.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Активация**

 переменная _expression_A, представляющий объект **OLEFormat** .


## <a name="remarks"></a>Заметки

Microsoft Publisher выполняется в одном окне, с помощью метода **активировать** с помощью объекта **Window** , поэтому Publisher активное приложение.


## <a name="example"></a>Пример

В следующем примере создается Publisher активное приложение.


```vb
Application.ActiveWindow.Activate
```

В следующем примере добавляет таблицы Excel в первой страницы публикации, активных и активирует электронной таблицы для редактирования.




```vb
Dim shpSheet As Shape 
 
Set shpSheet = ActiveDocument.Pages(1).Shapes.AddOLEObject (Left:=72, Top:=72, ClassName:="Excel.Sheet") 
 
shpSheet.OLEFormat.Activate
```


