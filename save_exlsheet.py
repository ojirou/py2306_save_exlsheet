import openpyxl
BaseFolder = input('複数のシートを含むエクセルファイルが入っているフォルダを入力　>> ')
if(BaseFolder[-1:]!="\\"):
    BaseFolder=BaseFolder + '\\'
FileName=input('Excelファイル名を入力 例：〇.xlsx >> ')
i=0
for i in range(0,10):
    wb=openpyxl.load_workbook(BaseFolder+FileName)
    try:
        ws_target=wb.worksheets[i].title
    except:
        break
    for ws in wb.worksheets:
        if ws.title == ws_target:
            continue
        else:
            wb.remove(ws)
        wb.save(BaseFolder + ws_target + ".xlsx")
        i=i+1