# import jinja2
# from openpyxl import *
# import pandas as pd

# wb = Workbook()
# ws = wb.active
# wb.save('мой.xlsx')
# # assign data of lists.
# data = {'Name': ['3', '5', '7', '6'], 'Age': [20, 21, 19, 18]}

# df = pd.DataFrame(data)

# print(df)
# a = df.style.applymap(lambda x: "background-color: red" if int(x)>12 else "background-color: white")


# a.to_excel('мой.xlsx',engine='openpyxl',index=False)