from win32com.client import Dispatch
import win32com.client
import os


class operatExcel:
    """Classes are used to create EXCEL,

     open EXCEL,edit EXCEL cell format, save EXCEL, close EXCEL
    Attributes:
        xlApp: Dispatch 'Excel.Application'
        filename: Absolute path to the excel document
        xlBook: open the excel document
    """

    def __init__(self, filename=None):
        """Open file or create new file (if it doesn't exist)"""
        self.xlApp = win32com.client.Dispatch('Excel.Application')
        if filename:
            self.filename = filename
            self.xlBook = self.xlApp.Workbooks.Open(filename)
        else:
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''

    def new(self,newfilename):
        """Create a new EXCEL file

        Args:
            newfilename:Created excel name
            """
        self.filename = newfilename
        self.xlBook.SaveAs(self.filename)

    def save(self, newfilename=None):
        """Save the EXCEL file

        Args:
            newfilename: saved excel name"""
        if newfilename:
            self.filename = newfilename
            self.xlBook.SaveAs(self.filename)
        else:
            self.xlBook.Save()

    def close(self):
        """Close the EXCEL file

        :return:
        """
        self.xlBook.Close(SaveChanges=0)
        del self.xlApp

    def getCell(self, sheet, row, col):
       """Get value of one cell

       Args:
            sheet: the EXCEL sheet name
            row: Row of EXCEL sheet
            col: Column of EXCEL sheet
       :return:sht.Cells(row, col).Value,the value of the cell"
       """
       sht = self.xlBook.Worksheets(sheet)
       return sht.Cells(row, col).Value

    def setCell(self, sheet, row, col, value):
        """set value of one cell

         Args:
            sheet: the EXCEL sheet name
            row: Row of EXCEL sheet
            col: Column of EXCEL sheet
            value: Set the value of the cell
        :return:
        """
        sht = self.xlBook.Worksheets(sheet)
        sht.Cells(row, col).Value = value

    def setCellformat(self, sheet, row, col):
        """Cell format adjustment

         Args:
            sheet: the EXCEL sheet name
            row: Row of EXCEL sheet
            col: Column of EXCEL sheet
        :return:
        """
        sht = self.xlBook.Worksheets(sheet)
        sht.Cells(row, col).Font.Size = 15  # font size
        sht.Cells(row, col).Font.Bold = True  # Whether it is black
        sht.Cells(row, col).Name = "Arial"  # Font type
        sht.Cells(row, col).Interior.ColorIndex = 3  # Form background
        # sht.Range("A1").Borders.LineStyle = xlDouble
        sht.Cells(row, col).BorderAround(1, 4)  # Table border
        sht.Rows(3).RowHeight = 30  # Row height
        sht.Cells(row, col).HorizontalAlignment = -4131  # Horizontally centered xlCenter
        sht.Cells(row, col).VerticalAlignment = -4160


if __name__ == "__main__":
    #测试在当前路径下新建一个mybook1的excel表格
    proDir = os.path.split(os.path.realpath(__file__))[0]
    path = os.path.join(proDir, "mybook5")
    xls = operatExcel()
    #新建一个excel
    xls.new(path)
    xls.close()
    #测试打开一个已经存在的excel文件
    xls = operatExcel(path)
    #测试设置单元格（2，A)的数值为10086
    xls.setCell('sheet1', 2, 'A', 10086)
    #测试获取单元格的数值
    cellvalue = xls.getCell('sheet1', 2, 'A')
    print(cellvalue)
    #测试对单元格的格式进行调整
    xls.setCellformat('sheet1',3,1)
    #保存文件
    xls.save()
    #关闭文件
    xls.close()

