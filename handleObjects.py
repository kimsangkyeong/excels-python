#############################################################################################################################
#
# Author : Sang kyeong Kim ( kimsangkyeong@gmail.com )
# Description : If Excel's Objects is used, it is an example program that automatically saves the values by reading them. 
#               Can be applied with the ability to automatically collect answers in questionnaires
# Information : https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook
#               https://pypi.org/project/pywin32/
#               install win32com.client  : pip install pywin32
#               install xlwt             : pip install xlwt
#               https://docs.microsoft.com/en-us/office/vba/api/excel.xldirection
#               Data를 포함하는 마지막 Row, Column Search를 위한 파라미터 : xlUp(-4162), xlDown(-4121), xlToLeft(-4159), xlToRight(-4161)
#               ex) ws.Range('B1').End(-4121).Row, ws.Range('A9').End(-4161).Column) 
# Dependency  : execel_objects.xlsx (Questionnaire file with predefined objects set)
#               pywin32 - release 228
#############################################################################################################################

from win32com.client import gencache # pip install pywin32
import xlwt                          # pip install xlwt
import os

#------< Questionair File Handling Functions >---
# display list items
def dsp_list(inputs):
  print("List count : %d" % len(inputs))
  for input in inputs:
    print(input)

# display worksheet's Shapes items
def dsp_shapes(shapes):
  print('Shape Count: %s' % len(shapes))
  for shape in shapes:
      # shape.Type => 8 : developer tools objects
      # shape.FormControlType => 1 : Check Box, 2 : Drop Down, 7 : Option Button
      # shape.Visible => -1 : visible,  0 : invisible
      # shape.ControlFormat.Value => 1 : Checked , -4146 : unChecked in Check Box, Option Button Object
      #                              0,1,2... : Selected Index in Drop Down
      if shape.FormControlType == 2:
        print('ListFillRange : %s' % shape.ControlFormat.ListFillRange)
        print('ListIndex : %s' % shape.ControlFormat.ListIndex)
        print('selected value : %s' % ws.Range(shape.ControlFormat.ListFillRange).Offset(shape.ControlFormat.ListIndex))
      print('shape Creator : %s, type : %s' % (shape.Creator, shape.Type))
      print('shape Name : %s' % shape.Name)
      print('shape ID : %s' % shape.ID)
      print('shape Visible : %s' % shape.Visible)
      print('FormControlType : %s' % shape.FormControlType)
#      print('Left %s, Top %s ' % (shape.Left, shape.Top))
      print('Parent %s, Placement %s' % (shape.Parent, shape.Placement))
#      print('ZOrderPosition %s' % shape.ZOrderPosition)
      print('Alternative Text : %s ' % shape.AlternativeText)
      print('ControlFormat.Value : %s' % shape.ControlFormat.Value)
      print('ControlFormat.LinkedCell : %s' % shape.ControlFormat.LinkedCell)
      print('ControlFormat.Parent : %s' % shape.ControlFormat.Parent)
#      print('shape TopLeftCell.Row: %s, Column : %s' % (shape.TopLeftCell.Row, shape.TopLeftCell.Column))
#      print('shape.BottomRightCell.Row :%s , Column : %s' % (shape.BottomRightCell.Row, shape.BottomRightCell.Column))
#      print('ControlFormat.Application.Cells : %s' % shape.ControlFormat.Application.Cells)
      print('.....')

# extract answer items in single file and append items to answers List
def extract_answer(qfilepath, answers, excel):
  # file read
  wb = excel.Workbooks.Open(qfilepath)  #wb = excel.Workbooks.open(r"c:\tmp\excel_objects.xlsx")
  ws = wb.Worksheets('Questionair') # ws = wb.Worksheets.Item(1)
  #print(ws.Name)
  #print('show all data : %s' % ws.ShowAllData )
  #print('row Count: %s' % ws.Rows.Count )
  #print('row Height: %s' % ws.Rows(1).RowHeight )
  #print('column Count: %s' % ws.Columns.Count )
  #print('cells.Row : %s, Cells.Column: %s' % (ws.Cells.Row, ws.Cells.Column) )
  #print('cells.Row : %s, Cells.Column: %s' % (ws.Rows.End(-4121).Row, ws.Columns.End(-4161).Column) )
  #print('cells.Row : %s, Cells.Column: %s' % (ws.Range('B1').End(-4121).Row, ws.Range('A9').End(-4161).Column) )
  # https://docs.microsoft.com/en-us/office/vba/api/excel.xldirection
  # Data를 포함하는 마지막 Row, Column Search를 위한 파라미터 : xlUp(-4162), xlDown(-4121), xlToLeft(-4159), xlToRight(-4161)
  
  
  # display list Shapes : develop tool's objects
  #dsp_shapes(ws.Shapes)
  
  # add all answer 
  for cell in ws.Range(ws.Cells(5,1), ws.Cells(5,4)):  # Set the number of lists to organize the survey subjects at once. (4 -> 50 ? )
    if cell.Column >= 3 :      # 3 - Survey Target Collection Start Column
      if cell.Text == '' :     # If no information is available, the End Survey Collections column
        break
  
      #default answer setting
      answer = {'colunminfo':'','surveyname':'', 'writer':'','wemail':'','q1':'','q2':'','q3-1':'','q3-2':'','q3-4':'','q4':'','q5':'','q6':'','q7':'','q8':''}
      
      # set writer info - common
      answer['writer'] = ws.Cells(4,2).Text
      answer['wemail'] = ws.Cells(5,2).Text
  
      # set survey info
      answer['surveyname'] = cell.Text
      answer['colunminfo'] = cell.Column
      answer['q1']         = ws.Cells(8, cell.Column).Text
      answer['q2']         = ws.Cells(9, cell.Column).Text
      answer['q4']         = ws.Cells(11, cell.Column).Text
      answer['q8']         = ws.Cells(15, cell.Column).Text
      for shape in ws.Shapes:
        if shape.Type == 8:
          prefix = str(cell.Column) + '_'
          q3_1str = prefix + 'chk_favor_1'
          q3_2str = prefix + 'chk_favor_2'
          q3_3str = prefix + 'chk_favor_3'
          q3_4str = prefix + 'chk_favor_4'
          q5str   = prefix + 'ddn_favornum'
          q6str   = prefix + 'ddn_hatenum'
          q7str   = prefix + 'opt_man'
          if q3_1str == shape.Name:
            answer['q3-1'] = 'Yes' if shape.ControlFormat.Value == 1 else 'No'
          if q3_2str == shape.Name:
            answer['q3-2'] = 'Yes' if shape.ControlFormat.Value == 1 else 'No'
          if q3_3str == shape.Name:
            answer['q3-3'] = 'Yes' if shape.ControlFormat.Value == 1 else 'No'
          if q3_4str == shape.Name:
            answer['q3-4'] = 'Yes' if shape.ControlFormat.Value == 1 else 'No'
          if q5str == shape.Name:
            if shape.ControlFormat.Value > 0:
              answer['q5'] = ws.Range(shape.ControlFormat.ListFillRange).Offset(shape.ControlFormat.ListIndex).Text
            else:
              answer['q5'] = '미선택'
          if q6str == shape.Name:
            answer['q6'] = ws.Range(shape.ControlFormat.ListFillRange).Offset(shape.ControlFormat.ListIndex).Text if shape.ControlFormat.Value > 0 else '미선택'
          if q7str == shape.Name:
            answer['q7'] = '남자' if shape.ControlFormat.Value == 1 else '여자'
              
          #print('shape Name : %s, %s' % (shape.Name, shape.ControlFormat.Value ))
      #print(answer)
      answers.append(answer)
  
  # display list data
  #dsp_list(answers)
  
  #workbook close
  wb.Close()

#------< Master File Handling Functions >---
# find the next row cursor for writing
def find_next_row_cursor(worksheet):
  for cell in worksheet.Range(worksheet.Cells(1,1), worksheet.Cells(worksheet.Rows.Count,1)):
    if cell.Row >= 2 :      # 3 - Survey Target Collection Start Column
      if cell.Text == '':
        return cell.Row
  return 2

# set header data and set write row cursor
def set_header_data(worksheet, headrow):
  worksheet.Cells(headrow,1).Value  = '순번'
  worksheet.Cells(headrow,2).Value  = '참조파일'
  worksheet.Cells(headrow,3).Value  = '참조파일Column'
  worksheet.Cells(headrow,4).Value  = '설문대상자'
  worksheet.Cells(headrow,5).Value  = '작성자'
  worksheet.Cells(headrow,6).Value  = '작성자이메일'
  worksheet.Cells(headrow,7).Value  = 'q1'
  worksheet.Cells(headrow,8).Value  = 'q2'
  worksheet.Cells(headrow,9).Value  = 'q3-1'
  worksheet.Cells(headrow,10).Value  = 'q3-2'
  worksheet.Cells(headrow,11).Value  = 'q3-3'
  worksheet.Cells(headrow,12).Value  = 'q3-4'
  worksheet.Cells(headrow,13).Value  = 'q4'
  worksheet.Cells(headrow,14).Value = 'q5'
  worksheet.Cells(headrow,15).Value = 'q6'
  worksheet.Cells(headrow,16).Value = 'q7'
  worksheet.Cells(headrow,17).Value = 'q8'

# write answers list items
def merge_list(inputs, qfile, worksheet, startrow):
#  print("merge List count : %d" % len(inputs))
  write_row_cursor = startrow
  for input in inputs:
    print(input)
    worksheet.Cells(write_row_cursor,1).Value = write_row_cursor - 1
    worksheet.Cells(write_row_cursor,2).Value = qfile
    worksheet.Cells(write_row_cursor,3).Value = input['colunminfo']
    worksheet.Cells(write_row_cursor,4).Value = input['surveyname']
    worksheet.Cells(write_row_cursor,5).Value = input['writer']
    worksheet.Cells(write_row_cursor,6).Value = input['wemail']
    worksheet.Cells(write_row_cursor,7).Value = input['q1']
    worksheet.Cells(write_row_cursor,8).Value = input['q2']
    worksheet.Cells(write_row_cursor,9).Value = input['q3-1']
    worksheet.Cells(write_row_cursor,10).Value = input['q3-2']
    worksheet.Cells(write_row_cursor,11).Value = input['q3-3']
    worksheet.Cells(write_row_cursor,12).Value = input['q3-4']
    worksheet.Cells(write_row_cursor,13).Value = input['q4']
    worksheet.Cells(write_row_cursor,14).Value = input['q5']
    worksheet.Cells(write_row_cursor,15).Value = input['q6']
    worksheet.Cells(write_row_cursor,16).Value = input['q7']
    worksheet.Cells(write_row_cursor,17).Value = input['q8']
    write_row_cursor += 1
  return write_row_cursor

#------< Main process >---
# Excel.Application Call
excel = gencache.EnsureDispatch('Excel.Application')
excel.Visible = False

# Open a file to marge your answers
master_fname = 'c:\\tmp\\AnswerMaster.xlsx' 
if os.path.exists(master_fname) : # exist file
  print('file exists')
  # Open Exist File
  mwb = excel.Workbooks.Open(master_fname) 
  mws = mwb.Worksheets('AnswerMaster') # ws = wb.Worksheets.Item(1)

  # find the next row cursor for writing
  write_row_cursor = find_next_row_cursor(mws)
else:  # new file
  print('file not exists')
  # Add new Workbook
  mwb = excel.Workbooks.Add()
  mws = mwb.Worksheets.Item(1)

   # set header data and set write row cursor
  set_header_data(mws, 1)
  write_row_cursor = 2


# exisitence of questionair file
exist_questionair = False

# search questionair directory
questionair_dir = 'c:\\tmp\\Questionair' 
for qfile in os.listdir(questionair_dir):
  qfilepath= questionair_dir + '\\' + qfile  # create full path file name
  if os.path.isfile(qfilepath) : # file type
    #print('file ', qfilepath)
    exist_questionair = True
    # write list data to merge file
    answers = list()
    extract_answer(qfilepath, answers, excel)
    # display list data
    #dsp_list(answers)
    write_row_cursor = merge_list(answers, qfile, mws, write_row_cursor)

# Only if the questionair file exists
if exist_questionair :
  if os.path.exists(master_fname) : # exist file
    mwb.Save()
  else:  # new file
    # create new file
    mws.Name = "AnswerMaster"
    mwb.SaveAs(master_fname)
  
#workbook close
mwb.Close()
#Application close
excel.Application.Quit()
