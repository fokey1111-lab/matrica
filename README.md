# RS Matrix Nasdaq Style

Готовый проект для репозитория.

## Что делает
- Загружает Excel `.xlsx` / `.xls`
- Берёт первый лист
- Ожидает структуру:
  - колонка `A` = `DateTime`
  - колонки `B..` = активы
- Строит weekly Friday closes
- Считает RS Matrix в логике P&F:
  - `BX`
  - `BO`
  - `SO`
  - `SX`
- Сортирует активы от слабых к сильным
- Показывает матрицу в визуальном стиле, близком к Nasdaq/Dorsey Wright
- Даёт экспорт матрицы в CSV

## Запуск
```bash
npm install
npm run dev
```

## Сборка
```bash
npm run build
```

## Пример файла
В папке `public` лежит `sample-data.xlsx`.
