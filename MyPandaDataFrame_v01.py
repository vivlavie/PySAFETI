import pandas as pd 
from openpyxl import load_workbook
iExl = pd.ExcelFile('v08_9mi.xlsx')
tlv = pd.read_excel(iExl,'Time varying leak')
#how to skip reading #1 ~ #61
#header row at #62

cExl = pd.ExcelFile('v08_9mc.xlsx')
dc = pd.read_excel(cExl,'Discharge')
jf = pd.read_excel(cExl,'Jet fire')
epf = pd.read_excel(cExl,'Early Pool Fire')
lpf = pd.read_excel(cExl,'Late Pool Fire')


not_a_tlv_index = tlv[tlv.iloc[:,0]!='Yes'].index
tlvonly = tlv.drop(not_a_tlv_index)
tlvonly.columns = tlvcolname

