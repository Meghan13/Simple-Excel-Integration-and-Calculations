import ExcelReaderWriter

xlrw = ExcelReaderWriter.ExcelReaderWriter("Data.xlsx")

print(xlrw.excel_reader((1,1)))
xlrw.excel_writer("did this work??", (2, 2))
