from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import openpyxl as pyxl
from .ScheduleUI import parse_args
"""
Program: ScheduleORganizer.py
Purpose: To automatically generate and format excel sheets as venue grids. Formatting
an Excel spreadsheet(s) with a visual grid & separate shift list, automated with Excel
formulas
"""

#HELPER FUNCTIONS
#Cells: [date,venue,position,time], Sheet: Sheet Obj, Ref_Sheet: Reference Sheet Title, Vol_Cell: Value of Volunteer Cell
def insert_sheet_list(cells,sheet,ref_sheet,vol_cell):
        sheet_call = ref_sheet.title
        space_formula = '&" "&'
        #Apostrophe & Spaces in Sheet Title Check w/ Replacement
        if sheet_call.find("'") != -1: sheet_call = sheet_call.replace("'","\'\'")
        if sheet_call.find(" ") != -1: sheet_call = "'"+sheet_call+"'"
        sheet_call = sheet_call+"!"
        #Finding inital cell for position cells, does the same to time cells if universal shift time is required
        position_merge = [r for r in ref_sheet.merged_cells.ranges if cells[2] in r][0].start_cell.coordinate
        if args.Universal_Shift_Time == True: time_value = [r for r in ref_sheet.merged_cells.ranges if cells[3] in r][0].start_cell.coordinate
        else: time_value = cells[3]
        #Formatting excel formula & inserting into shift list sheet
        #Check if next cell is empty in name column, inserts name & shift info in adjacent cell-column
        r=1
        while True:
                r +=1
                cell_name = sheet.cell(row=r,column=1)
                create_cell_border(cell_name)
                if cell_name.value is None: 
                        cell_name.value = "="+sheet_call+vol_cell.coordinate
                        cell_shift = sheet.cell(row = cell_name.row, column = cell_name.column+1)
                        cell_shift.value = "="+sheet_call+str(cells[0])+space_formula+sheet_call+str(cells[1])+space_formula+sheet_call+str(position_merge)+space_formula+sheet_call+str(time_value)
                        create_cell_border(cell_shift)
                        break
#Creates cell borders, given excel cell
def create_cell_border(cell):
        thin_border = Border(left=Side(style='thin'), 
                     right=Side(style='thin'), 
                     top=Side(style='thin'), 
                     bottom=Side(style='thin'))
        cell.border = thin_border
#Returns either 0 or 1, if a given cell is within the list of merged cells in a given sheet
def is_merged(cell,sheet):
        merged = 0
        for mergedCell in sheet.merged_cells.ranges:
                if cell.coordinate in mergedCell:
                        merged = 1
                        break
        return merged

#MAIN FUNCTION
if __name__ == '__main__':
    print("Creating Venue Schedule Grid")
    args = parse_args()
    filename = args.File_Name
    wb = pyxl.load_workbook(filename, keep_vba=True)
    wb.save(filename)
    if args.Venue_Name in wb.sheetnames: 
        print('Venue Name already used in spreadsheet, exiting generation process')
        exit()
    else:
        wb.create_sheet(args.Venue_Name)
        sheet = wb[args.Venue_Name]
    if "ShiftList" in wb.sheetnames: sheet_list = wb['ShiftList']
    else:
        wb.create_sheet("ShiftList")
        sheet_list = wb['ShiftList']
    print(args.Split_Shifts)
    wb.save(filename)
    pos_vols = [args.BG_Vols,args.FOH_Vols,args.GT_Vols,args.Hosp_Vols,args.Merch_Vols,args.Stage_Vols,args.Sec_Vols]
    pos_names = ['Beer Garden','Front of House','Green Team','Hospitality','Merchandise','Staging','Security']
    #General Font & Alignment for venue grid
    font = Font(name='Calibri',
                size=28,
                bold=True,
                italic=False,
                underline='none',
                strike=False,
                color='FF000000')
    align = Alignment(horizontal='center',
                    vertical='bottom',
                    text_rotation=0,
                    wrap_text=True,
                    shrink_to_fit=False,
                    indent=0)
    #Using Column A as base cells, for chart formatting i.e. all formatting for merging and insertion will be done in column A
    #Make Venue Name & Shift List Headers
    sheet['A1'] = args.Venue_Name
    sheet['A1'].font = font
    sheet['A1'].alignment = align
    sheet_list['A1'] = 'Volunteers'
    sheet_list['B1'] = 'Shifts'
    create_cell_border(sheet['A1'])
    sheet.column_dimensions['A'].width = len(args.Venue_Name) #Figure out good cell width value
    sheet.row_dimensions[1].height = 30
    sheet.merge_cells(start_row=1, start_column=1,end_row=1,end_column=int(args.Number_of_Shows))
    sheet['A1'].fill = PatternFill("solid", fgColor="FFC3A8")
    start = 4
    for i in range(int(args.Number_of_Shows)):
        #Counter keeps track of cell traversals in real time, for each column, resetting once a new column is selected
        counter = 0
        #pos_index is an internal counter for pos_names
        pos_index = 0
        #Create Date Cell
        cell = sheet.cell(row=2,column=1+i)
        create_cell_border(cell)
        cell_date = cell.coordinate
        #Create Show Title Cell
        cell = sheet.cell(row=3,column=1+i)
        create_cell_border(cell)
        #For every volunteer in a position, create a position, supervisor, and volunteer cells
        for vols in pos_vols:
                if int(vols) !=0:
                        #Position Name Cell
                        cell = sheet.cell(row=start+counter,column=1+i)
                        merged = is_merged(cell,sheet)
                        if merged == 0:
                                cell = sheet.cell(row=start+counter,column=1+i)
                                create_cell_border(cell)
                                cell.value = pos_names[pos_index] #Inserting Volunteer Position into cell
                                sheet.merge_cells(start_row=start+counter, start_column=1,end_row=start+counter,end_column=int(args.Number_of_Shows))
                                sheet.cell(row=start+counter,column=1+i).alignment = align
                                cell_position = cell.coordinate
                        else: cell_position = cell.coordinate
                        split_shift = True
                        counter += 1
                        while True:
                                #Shift Time Cell
                                cell = sheet.cell(row=start+counter,column=1+i)
                                merged = is_merged(cell,sheet)
                                #Shift Time Cell Merge Requirement Check
                                if args.Universal_Shift_Time == True and merged == 0:
                                        create_cell_border(cell)
                                        cell.value = args.Shift_Time
                                        sheet.merge_cells(start_row=start+counter, start_column=1,end_row=start+counter,end_column=int(args.Number_of_Shows))
                                if args.Universal_Shift_Time == None:
                                        create_cell_border(cell)
                                        cell.value = args.Shift_Time
                                sheet.cell(row=start+counter,column=1+i).alignment = align
                                cell_time = cell.coordinate
                                #Supervisor Insertion
                                if pos_names[pos_index] in args.Supervisor_Positions: 
                                        sheet.cell(row=start+counter+1,column=1+i,value = pos_names[pos_index] + ' Supervisor')
                                        create_cell_border(sheet.cell(row=start+counter+1,column=1+i))
                                        counter +=1
                                cell_venue = sheet.cell(row=1,column=1)
                                #Volunteer Insertion
                                for j in range(int(vols)): 
                                        vol_cell = sheet.cell(row=start+counter+1+j,column=1+i,value = 'Test')
                                        insert_sheet_list([cell_date,cell_venue.coordinate,cell_position,cell_time],sheet_list,sheet,vol_cell)
                                        wb.save(filename)
                                        create_cell_border(sheet.cell(row=start+counter+1+j,column=1+i))
                                counter +=int(vols)+1
                                #Split Shift Check, switching boolean flag to break on repeat. Breaks if None otherwise
                                if split_shift is False: break
                                if args.Split_Shifts is None: break
                                if pos_names[pos_index] in args.Split_Shifts: split_shift = False
                                else: break
                pos_index +=1
    #Adjusting column width of entire sheet, based on maximum cell value length
    #Taking filled active cell values maximum value & adjusting global column widths
    dims = {}
    for row in sheet.rows:
        for cell in row:
                if cell.value: dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))  
    for col, value in dims.items(): sheet.column_dimensions[col].width = value
    wb.save(filename)
    print("Done")
