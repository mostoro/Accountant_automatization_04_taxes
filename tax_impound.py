# MOISES TORO
#PROPER AI


#Importing libraries
import xlwings as xw

#Reading schedules
user = input("Enter your laptop user: ")
mec = input("Enter Current Closing Month (Example for March:'03'):")


file_to_update_1 = xw.Book(r'C:..file_1{file}.xlsx'.format(folder=mec,file=mec))
file_to_update_2 = xw.Book(r'C:..file_2{file}.xlsx'.format(folder=mec,file=mec))
file_to_update_3 = xw.Book(r'C:..file_3{file}.xlsx'.format(folder=mec,file=mec))
file_to_update_4 = xw.Book(r'C:..file_4{file}.xlsx'.format(folder=mec,file=mec))
file_to_update_5 = xw.Book(r'C:..file_5{file}.xlsx'.format(folder=mec,file=mec))
file_to_update_6 = xw.Book(r'C:..file_6{file}.xlsx'.format(folder=mec,file=mec))
file_to_update_7 = xw.Book(r'C:..file_7{file}.xlsx'.format(folder=mec,file=mec))
file_to_update_8 = xw.Book(r'C:..file_8{file}.xlsx'.format(folder=mec,file=mec))
file_to_update_9 = xw.Book(r'C:..file_9{file}.xlsx'.format(folder=mec,file=mec))



files_to_update = [file_to_update_1, file_to_update_2, file_to_update_3, file_to_update_4
                   , file_to_update_5, file_to_update_6, file_to_update_7, file_to_update_8, file_to_update_9]




# cpi_i = xw.Book(r'2023-{file} - CPI I - Tax Impound.xlsx')
# cpi_ii = xw.Book(r'2023-{file} - CPI II - Tax Impound.xlsx')
# bmf1 = xw.Book(r'2023-{file} - BMF I - Tax Impound.xlsx')
# bmf2 = xw.Book(r'2023-{file} - BMF II - Tax Impound.xlsx')
# sfmf1 = xw.Book(r'2023-{file} - SFMF I - Tax Impound.xlsx') 
# sfmf2 = xw.Book(r'2023-{file} - SFMF II - Tax Impound.xlsx')
# sfmf3 = xw.Book(r'2023-{file} - SFMF III - Tax Impound.xlsx') 




# files_to_update = [cpi_i, cpi_ii, bmf1, bmf2, sfmf1, sfmf2, sfmf3]



#Reading Yardi report
source_file = xw.Book(r'source_file.xlsx')
source = source_file.sheets.active  # in specific book


#Updating every schedule in files_to_update
cell_h = source['h1']
cell_i = source['i1']



for i in range(len(files_to_update)):

    current_schedule = files_to_update[i]

    sheet_names = [sheet.name for sheet in current_schedule.sheets]


    #Updating date for all properties

    for a_cell in sheet_names:
        if  a_cell != 'Revision':
            current_schedule.sheets['{:}'.format(a_cell)]['B4'].value = source['D8'].value


    #Updating schedule keys        
    schedule_keys = []

    for a_cell in source["A6"].expand("down"):
        if a_cell.value in sheet_names and (cell_i.offset(a_cell.row-1).value != 0 or cell_h.offset(a_cell.row-1).value != 0):
            for i in range(1,current_schedule.sheets[a_cell.value].range('b' + str(current_schedule.sheets[a_cell.value].cells.last_cell.row)).end('up').row):
                if current_schedule.sheets[a_cell.value]['B{}'.format(i)].value == '2023 Beginning Balance':
                    first_cell_date = i
                    break

            for i in range(current_schedule.sheets[a_cell.value]['b{}'.format(first_cell_date)].end('down').row -  current_schedule.sheets[a_cell.value]['b{}'.format(first_cell_date)].row + 1):
                schedule_keys.append(str(a_cell.value)+str(current_schedule.sheets[a_cell.value]['b{}'.format(first_cell_date + i)].value) + str(current_schedule.sheets[a_cell.value]['c{}'.format(first_cell_date + i)].value) + str(current_schedule.sheets[a_cell.value]['d{}'.format(first_cell_date + i)].value))


    #Updating schedules
    for a_cell in source["A6"].expand("down"):
        if a_cell.value in sheet_names and cell_h.offset(a_cell.row-1).value != 0:
            if not str(a_cell.value)+str(source['d1'](a_cell.row).value) + str(source['k1'](a_cell.row).value) + str(cell_h(a_cell.row).value ) in schedule_keys:
                

                current_schedule.sheets[a_cell.value].range('{}:{}'.format(current_schedule.sheets[a_cell.value]['d7'].end('down').row+1, current_schedule.sheets[a_cell.value]['d7'].end('down').row+1)).insert('down') #insert a row
                current_schedule.sheets[a_cell.value]['D{}'.format(current_schedule.sheets[a_cell.value]['d7'].end('down').row + 1)].value = cell_h.offset(a_cell.row-1).value 
                current_schedule.sheets[a_cell.value]['B{}'.format(current_schedule.sheets[a_cell.value]['d7'].end('down').row)].value = source['D1'].offset(a_cell.row-1).value
                current_schedule.sheets[a_cell.value]['C{}'.format(current_schedule.sheets[a_cell.value]['d7'].end('down').row)].value = source['K1'].offset(a_cell.row-1).value
                current_schedule.sheets[a_cell.value]['E{}'.format(current_schedule.sheets[a_cell.value]['d7'].end('down').row)].value = float(current_schedule.sheets[a_cell.value]['E{}'.format(current_schedule.sheets[a_cell.value]['d7'].end('down').row-1)].value) + float(current_schedule.sheets[a_cell.value]['D{}'.format(current_schedule.sheets[a_cell.value]['d7'].end('down').row)].value)

        if a_cell.value in sheet_names and cell_i.offset(a_cell.row-1).value != 0:
            if not str(a_cell.value)+str(source['d1'](a_cell.row).value) + str(source['k1'](a_cell.row).value) + str(cell_i(a_cell.row).value *-1) in schedule_keys:
                

                current_schedule.sheets[a_cell.value].range('{}:{}'.format(current_schedule.sheets[a_cell.value]['d7'].end('down').row+1, current_schedule.sheets[a_cell.value]['d7'].end('down').row+1)).insert('down') #insert a row
                current_schedule.sheets[a_cell.value]['D{}'.format(current_schedule.sheets[a_cell.value]['d7'].end('down').row + 1)].value = cell_i.offset(a_cell.row-1).value * -1
                current_schedule.sheets[a_cell.value]['B{}'.format(current_schedule.sheets[a_cell.value]['d7'].end('down').row)].value = source['D1'].offset(a_cell.row-1).value 
                current_schedule.sheets[a_cell.value]['C{}'.format(current_schedule.sheets[a_cell.value]['d7'].end('down').row)].value = source['K1'].offset(a_cell.row-1).value 
                current_schedule.sheets[a_cell.value]['E{}'.format(current_schedule.sheets[a_cell.value]['d7'].end('down').row)].value = float(current_schedule.sheets[a_cell.value]['E{}'.format(current_schedule.sheets[a_cell.value]['d7'].end('down').row-1)].value) + float(current_schedule.sheets[a_cell.value]['D{}'.format(current_schedule.sheets[a_cell.value]['d7'].end('down').row)].value)

    print(str(current_schedule),'updated')


