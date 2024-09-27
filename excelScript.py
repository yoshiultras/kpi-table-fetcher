import openpyxl
import database


def make_excel():
    # Получаем метрики из базы данных
    data = database.Database.get_metrics()
    # Создаем книгу и заполняем ее
    wb = openpyxl.Workbook()

    # добавляем новый лист
    wb.create_sheet(title='Метрики', index=0)

    # получаем лист, с которым будем работать
    sheet = wb['Метрики']
    ws = wb.active
    for row in range(len(data)):
        for col in range(1, len(data[row])):
            cell = sheet.cell(row=row + 1, column=col)
            # Для пропуска записи в недоступную ячейку после слияния
            try:
                cell.value = data[row][col]
            except Exception as err:
                continue
        # слияние ячеек с номерам критериев
        if row + 1 < len(data):
            if data[row][1] == data[row + 1][1]:
                ws.merge_cells(start_row=row + 1, start_column=1, end_row=row + 2, end_column=1)
                # Выставление ширины колонок
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 5
    ws.column_dimensions['C'].width = 80
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 30
    ws.column_dimensions['F'].width = 30
    ws.column_dimensions['G'].width = 30
    ws.column_dimensions['H'].width = 25
    ws.column_dimensions['I'].width = 70
    ws.column_dimensions['J'].width = 70
    ws.column_dimensions['K'].width = 15

    wb.save('example.xlsx')
    print("Файл успешно сохранен")


make_excel()
