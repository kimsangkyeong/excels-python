from win32com.client import gencache # pip install pywin32
import xlwt   # pip install xlwt

# https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook

# file read
dest_filename = 'c:\\tmp\\sample.xlsx' 
excel = gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(dest_filename)
#wb = excel.Workbooks.open(r"c:\tmp\sample.xlsx")
#excel.Visible = True

# ws = wb.Worksheets.Item(1)
print('......................>>> ....')
ws = wb.Worksheets('ksheet')
#print('show all data : %s' % ws.ShowAllData )
print('row Count: %s' % ws.Rows.Count )
print('row Height: %s' % ws.Rows(1).RowHeight )
print('column Count: %s' % ws.Columns.Count )
print('cells.Row : %s, Cells.Column: %s' % (ws.Cells.Row, ws.Cells.Column) )
print('cells.Row : %s, Cells.Column: %s' % (ws.Rows.End(-4121).Row, ws.Columns.End(-4161).Column) )
print('cells.Row : %s, Cells.Column: %s' % (ws.Range('B1').End(-4121).Row, ws.Range('A9').End(-4161).Column) )
# https://docs.microsoft.com/en-us/office/vba/api/excel.xldirection
# Data를 포함하는 마지막 Row, Column Search를 위한 파라미터 : xlUp(-4162), xlDown(-4121), xlToLeft(-4159), xlToRight(-4161)
print('Shape Count: %s' % len(ws.Shapes))
for shape in ws.Shapes:
  if shape.Type == 8: # form control
    if 'chk11' in shape.Name:
      print('shape Creator : %s, type : %s' % (shape.Creator, shape.Type))
      print('shape Name : %s' % shape.Name)
      print('shape ID : %s' % shape.ID)
      print('shape Visible : %s' % shape.Visible)
      print('FormControlType %s, Left %s, Top %s ' % (shape.FormControlType, shape.Left, shape.Top))
      print('Parent %s, Placement %s' % (shape.Parent, shape.Placement))
      print('ZOrderPosition %s' % shape.ZOrderPosition)
      print('Alternative Text %s: ControlFormat.Value %s' % (shape.AlternativeText, shape.ControlFormat.Value))
      print('ControlFormat.LinkedCell : %s' % shape.ControlFormat.LinkedCell)
      print('ControlFormat.Parent : %s' % shape.ControlFormat.Parent)
      print('shape TopLeftCell.Row: %s, Column : %s' % (shape.TopLeftCell.Row, shape.TopLeftCell.Column))
      print('shape.BottomRightCell.Row :%s , Column : %s' % (shape.BottomRightCell.Row, shape.BottomRightCell.Column))
#    print('ControlFormat.Application.Cells : %s' % shape.ControlFormat.Application.Cells)

print(ws.Name)
print(ws.Range('B5'))

wb.Close()


# file write
fname = 'c:\\tmp\\sample3.xlsx'
#wb = excel.Workbooks.Add()
#ws = wb.Worksheets.Add()
#ws.Name = "MyNewSheet"
#wb.SaveAs(fname)

excel.Application.Quit()

