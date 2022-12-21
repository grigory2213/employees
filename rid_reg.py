import pandas as pd
import numpy as np

clients = pd.read_excel('/home/grigory/Local/Work/employees/Клиенты.xlsx')
rid = pd.read_csv('/home/grigory/Local/Work/employees/reg.csv')
surveys = pd.read_csv('/home/grigory/Local/Work/employees/Surveys.csv')
surveys_MY = pd.read_csv('/home/grigory/Local/Work/employees/surveys_MY.csv')

surveys = surveys[['SurveyInstanceID','WorkflowStatus','Client', 'Login', 'Campaign', 'SurveyStatusName','CP001','CP002','CP003','CP004','CP005','CP006','CP007','CP008','CP009']]

surveys_MY = surveys_MY[['ID','Клиент','Корректор','Создатель','Оплата клиента за визит','Оплата тайному покупателю (заложено)','Оплата корректору (заложено)','Руководитель проекта','Оплата руководителя проекта','Оплата координатора проекта','Другие расходы','Программное обеспечение','Прибыль']]
rid = rid[['SurveyInstanceID', 'Last Assigned', 'FirstValidator']]

surveys = surveys.loc[surveys['Login'] != 'Test']
surveys = surveys.loc[surveys['SurveyStatusName'] != 'Assigned - In "Working" status']
surveys = surveys.loc[surveys['SurveyStatusName'] != 'Assigned (Accepted where Acceptance is Required)']
surveys = surveys.loc[surveys['SurveyStatusName'] != 'Validation - Pending']
surveys = surveys.loc[surveys['SurveyStatusName'] != 'Validation - In Progress']
surveys = surveys.loc[(surveys['Campaign'] != '_Удаленные')]

merged_data = pd.merge(surveys, rid, how = 'left', left_on='SurveyInstanceID', right_on='SurveyInstanceID')

print(merged_data.info())

merged_data = merged_data.dropna(subset=['CP001','CP002','CP003','CP004','CP005','CP006','CP007','CP008','CP009'])
merged_data =merged_data.fillna('--')

print(merged_data)

merged_data = merged_data.rename({'SurveyInstanceID': 'ID','WorkflowStatus': 'status','Client' : 'Клиент', 'CP001': 'Оплата клиента за визит','CP002': 'Оплата тайному покупателю (заложено)','CP003': 'Оплата корректору (заложено)','CP004': 'Руководитель проекта','CP005':'Оплата руководителя проекта','CP006':'Оплата координатора проекта','CP007':'Другие расходы','CP008':'Программное обеспечение','CP009':'Прибыль','Last Assigned': 'Создатель', 'FirstValidator':'Корректор'}, axis=1)

print(merged_data)
print(merged_data.info())

merged_data = pd.concat([merged_data, surveys_MY], ignore_index=True)

merged_data['count'] = 1
print(merged_data)

merged_data_reg_number = pd.pivot_table(merged_data, values='count',index=['Создатель'], aggfunc=np.sum)

merged_data_reg_summ = pd.pivot_table(merged_data, values='Оплата координатора проекта',index=['Создатель'], aggfunc=np.sum)


merged_data_ruk_number = pd.pivot_table(merged_data, values='count',index=['Руководитель проекта'], aggfunc=np.sum)

merged_data_ruk_summ = pd.pivot_table(merged_data,values='Оплата руководителя проекта',index=['Руководитель проекта'], aggfunc=np.sum)


merged_data_rid_number = pd.pivot_table(merged_data, values='count',index=['Корректор'], aggfunc=np.sum)

merged_data_rid_summ = pd.pivot_table(merged_data, values='Оплата корректору (заложено)',index=['Корректор'], aggfunc=np.sum)





# merged_data_reg = pd.merge(merged_data_reg, clients, how='left', left_on='Client', right_on='Клиент')
# merged_data_reg = merged_data_reg.drop(['Клиент', 'Бюджет план', 'Ридер', 'Менеджер', 'Прибыль', 'Другие расходы', 'ТП', 'Остаток'], axis = 1)
# merged_data_reg['Итого за клиента'] = merged_data_reg['count'] * merged_data_reg['Координатор']

# merged_data_rid = pd.merge(merged_data_rid, clients, how='left', left_on='Client', right_on='Клиент')
# merged_data_rid = merged_data_rid.drop(['Клиент', 'Бюджет план', 'Координатор', 'Менеджер', 'Прибыль', 'Другие расходы', 'ТП', 'Остаток'], axis = 1)
# merged_data_rid['Итого за клиента'] = merged_data_rid['count'] * merged_data_rid['Ридер']

# merged_data_reg = merged_data_reg.rename(columns={'Last Assigned': 'Координатор', 'Client':'Клиент', 'count':'Кол-во', 'Координатор': 'Стоимость единицы'})

# merged_data_rid = merged_data_rid.rename(columns={'FirstValidator': 'Ридер', 'Client':'Клиент', 'count':'Кол-во', 'Ридер': 'Стоимость единицы'})

# pivot_table_regs = pd.pivot_table(merged_data_reg, values='Итого за клиента',index=['Координатор'], aggfunc=np.sum)

# pivot_table_rids = pd.pivot_table(merged_data_rid, values='Итого за клиента',index=['Ридер'], aggfunc=np.sum)

# print(merged_data_reg)
# print(merged_data_rid)


with pd.ExcelWriter('regs-rids.xlsx') as writer:  
    merged_data_reg_number.to_excel(writer, sheet_name='Координаторы_сводная_кол')
    merged_data_reg_summ.to_excel(writer, sheet_name='Координаторы_сводная_суммы')
    merged_data_ruk_number.to_excel(writer, sheet_name='Руководители_сводная_кол')
    merged_data_ruk_summ.to_excel(writer, sheet_name='Руководители_сводная_суммы')
    merged_data_rid_number.to_excel(writer, sheet_name='Ридеры_сводная_кол')
    merged_data_rid_summ.to_excel(writer, sheet_name='Ридеры_сводная_суммы')

# unique = list[merged_data_reg['Client'].unique()]
# print(unique)