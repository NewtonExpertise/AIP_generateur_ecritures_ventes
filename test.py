import xlwings as xw
import pandas as pd
import pprint

pp = pprint.PrettyPrinter(indent=4)


# def get_start_cell(file):

wb = xw.Book(r'V:\Mathieu\ventes_aip\test.xlsx')
# ws1 = wb.sheets[0]
# for ws in wb.sheets:
ws = wb.sheets['THC']
# print(ws.name)
# print(ws.range((1,1) ,(5,21)).value)


# for row in range(1,15):
#     for col in range(1,10):
#         if ws.range(row,col).value == "N° FACT":
            # print("col "+str(col)+"row "+str(row))
            # datas = ws.range(row,col).address
            # print(datas)
            # print(ws.range((row,col)).expand().value)
nbligne = ws.cells(ws.api.rows.count, "A").end(-4162).row
print(nbligne)
            # datas = ws.range((row,col),(str(nbligne), row+30)).value




# print(ws.range('A1').expand('right').value)
# # # # # nbligne = ws.cells(ws.api.rows.count, "A").end(-4162).row
# # # # # datas = ws.range('A6:X'+str(nbligne)).value
# # # # # for data in datas:
# # # # #     print(data)
# # print(ws.range())
# for row in range(1,15):
#     for col in range(1,15):
#         if ws.range(row,col).value == "Mathieu":
#             print(ws.range(ws.range(row,col).address).value)
#             print("row "+str(row)+"col "+str(col))
# df = pd.DataFrame([[1,2], [3,4]], columns=['a', 'b'])
# # ws.range('A1').value = df
# print(ws.range('A1').options(pd.DataFrame, expand='table').value[1])



# if __name__ == "__main__":
    
#     fichierexcel = r'V:\Mathieu\ventes_aip\test.xlsx'
#     x = get_start_cell(fichierexcel)
#     print(x)