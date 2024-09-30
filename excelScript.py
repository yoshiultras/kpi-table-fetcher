import openpyxl
import openpyxl.styles
import database

def set_border(ws, cell_range, need_to_thick, need_to_thick_up, need_to_thick_down):
    thin = openpyxl.styles.Side(border_style="thin", color="000000")
    thick = openpyxl.styles.Side(border_style="thick", color="000000")
    for row in ws[cell_range]:
        for cell in row:
                if cell == row[len(row)-1]:
                    cell.border = openpyxl.styles.Border(top=thin, left=thin, right=thick, bottom=thin)
                    cell.fill = openpyxl.styles.PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                else:
                    cell.border = openpyxl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
        if(len(need_to_thick) != 0):
            if row == ws[need_to_thick[0]]:
                for cell in row:
                    cell.border = openpyxl.styles.Border(top=thick, left=thick, right=thick, bottom=thick)    
                need_to_thick.pop(0)
        if(len(need_to_thick_up) != 0):
            if row == ws[need_to_thick_up[0]]:
                for cell in row:
                    if cell == row[len(row)-1]:
                        cell.border = openpyxl.styles.Border(top=thick, left=thin, right=thick, bottom=thin)
                    else:
                        cell.border = openpyxl.styles.Border(top=thick, left=thin, right=thin, bottom=thin)
                need_to_thick_up.pop(0)
        if(len(need_to_thick_down) != 0):
            if row == ws[need_to_thick_down[0]]:
                for cell in row:
                    if cell == row[len(row)-1]:
                        cell.border = openpyxl.styles.Border(top=thin, left=thin, right=thick, bottom=thick)
                    else:
                        cell.border = openpyxl.styles.Border(top=thin, left=thin, right=thin, bottom=thick)
                    
                need_to_thick_down.pop(0)
        
            

    
    

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
    
    # формирование шапки таблицы
    header_data = ["№","№","Показатели","Единица измерения", "Базовый уровень (минимальный) (k=1)", "Нормальный уровень (зона актуального развития) (k=1,5)", "Целевой уровень (зона ближайшего развития) (k=2)", "Периодичность измерения", "Условия оформления показателя", "Примечание", "Баллы"]
    ws['A1'].value = 'Приложение к Положению "О системе оценки показателей эффективности работы работников из числа заведующих кафедрами федерального государственного автономного образовательного учреждения высшего образования «Московский политехнический университет»'
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=11)
    ws['A1'].alignment = openpyxl.styles.Alignment(horizontal='center')
    ws['A2'].value = 'Перечень ключевых показателей эффективности деятельности заведующих кафедрами Московского Политеха (рассмотрен Ученым советом университета, протокол'
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=11)
    ws['A2'].alignment = openpyxl.styles.Alignment(horizontal='center')
    for i in range(len(header_data)):
        ws['E3'].value = 'Значение показателя (множитель)'
        ws['E3'].alignment = openpyxl.styles.Alignment(horizontal='center')
        ws.merge_cells(start_row=3, start_column=5, end_row=3, end_column=7)
        if(i <=3) or (7 <= i):
            cell = sheet.cell(row=3, column=i+1)
            cell.value = header_data[i]
            cell.alignment = openpyxl.styles.Alignment(wrap_text=True, horizontal='center', vertical='center')
            ws.merge_cells(start_row=3, start_column=i+1, end_row=4, end_column=i+1)
        else:
            cell = sheet.cell(row=4, column=i+1)
            cell.value = header_data[i]
            cell.alignment = openpyxl.styles.Alignment(wrap_text=True, horizontal='center', vertical='center')
            
    sections = database.Database.get_sections()
    curent_row = 5
    curent_category = 0
    counter=0
    need_to_thick = []
    need_to_thick_up = [3]
    need_to_thick_down = []
    for row in range(len(data)):
        #вставка категорий
        if(curent_category != int(data[row][11])):
                ws.insert_rows(curent_row)
                curent_category +=1
                sheet.cell(row=curent_row, column=1).value=sections[curent_category-1][1]
                ws.merge_cells(start_column= 1, start_row=curent_row, end_column=11, end_row=curent_row)
                need_to_thick.append(str(curent_row))
                curent_row+=1
        #Заполнение строк
        for col in range(len(data[row])-1):
            cell = sheet.cell(row=curent_row, column=col+1)
            cell.alignment = openpyxl.styles.Alignment(wrap_text=True, horizontal='center', vertical='center')
            # Для пропуска записи в недоступную ячейку после слияния
            try:
                cell.value = data[row][col]
            except Exception as err:
                continue
        # слияние ячеек с номерам критериев
        if row + 1 < len(data):
            if data[row][0] == data[row + 1][0]:
                counter+=1
                if counter ==1:
                    need_to_thick_up.append(curent_row)
            else:
                
                if counter==0 and data[row][1] == None:
                    ws.merge_cells(start_row=curent_row, start_column=1, end_row=curent_row, end_column=2)
                else:
                    ws.merge_cells(start_row=curent_row-counter, start_column=1, end_row=curent_row, end_column=1)
                    counter=0
                    need_to_thick_down.append(curent_row)
        else:
            ws.merge_cells(start_row=curent_row, start_column=1, end_row=curent_row, end_column=2)
            need_to_thick_down.append(curent_row)

        
        curent_row+=1

 
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
    set_border(ws, 'A3:K'+str(curent_row-1), need_to_thick, need_to_thick_up, need_to_thick_down) 
#Сохранение файла
    wb.save('Метрики.xlsx')
    print("Файл успешно сохранен")
 
        

make_excel()
