# -*- coding: utf-8 -*-
"""
Created on Thu Dec 20 14:24:20 2018

@author: HEMTP
"""
!pip install qgrid

import qgrid

qgrid_widget = qgrid.show_grid(ERD1, show_toolbar=True)
qgrid_widget 

!pip install fuzz
#------------------------------------------

import feather
import pandas as pd
import fuzz

#dframe = pd.read_excel("erd.xlsx")
Zcop1  = pd.read_excel("d:\\base.xlsx", sheet_name="Zcop")
compare_col = pd.read_excel("d:\\base.xlsx", sheet_name="Zcop")
ERD1 = pd.read_excel("d:\\base.xlsx", sheet_name="ERD")
#Indent1 = pd.read_excel("d:\\base.xlsx", sheet_name="Indent")
#Zcop1.head()
#ERD1.head()
#Indent1.head()

Template_columns = compare_col.columns
Zcop1_columns = Zcop1.columns

common_columns = Template_columns.intersection(Zcop1_columns)
diff_columns = Template_columns.difference(Zcop1_columns)                                               

common_columns
display(diff_columns)

ERD1['ASSIGN_END'] = ERD1.ASSIGN_END.astype(str)

ERD1.rename(columns={'EMP_NO':'EMP_CODE', 'ASSIGN_END':'ERD_DATE'}, inplace=True)

ERD2 = ERD1[['EMP_CODE', 'ERD_DATE']]

ERD2.dtypes

##LIST COLUMNS
#ALL_COLUMNS=list(ERD3)
#ALL_COLUMNS

#ERD3.rename(columns={'EMP_NO':'EMP_CODE', 'ASSIGN_END':'ERD_DATE'}, inplace=True)

#ERD_dict = ERD2.to_dict()


#ERD_dict

#Zcop2 = Zcop1.merge(ERD2, on='EMP_NO')

ERD4 = ERD2.set_index('EMP_CODE').T.to_dict('list')

ERD4
#ERD2.set_index("EMP_CODE", drop=True, inplace=True)

#ERD3 = ERD2.to_dict(orient="index")

Zcop1['ERD_Date'] = Zcop1.EMP_CODE.map(ERD4)

 
#types_dict = Zcop1.dtypes.to_dict()
#types_dict

#cols = Zcop1.columns.tolist()
#cols
#print(os.getcwd())
#cols = list(Zcop1.columns.values)
Zcop2 = Zcop1[['LOAD_DATE',	'COM_CODE',	'COSTCTR',	'CC_DESC',	'SAP_BU_DESC',	'SAP_VERTICAL_DESC',	'CUST_NUM',	'CUST_NAME',	'PROJ_TYPE',	'WBS_CODE',	'DOM_ID',	'DOM_DESC',	'HOME_ORGUNIT',	'HOME_ORGUNIT_DESC',	'ORG_UNIT',	'ORG_UNIT_DESCRIPTION',	'PROJECT_CODE',	'STAFFING_PERCENTAGE',	'BILLABILITY_STATUS',	'EMP_CODE',	'EMP_NAME',	'PM_ID',	'PM_NAME',	'TM_EMP_NO',	'TM_NAME',	'DERIVED_EMP_CITY',	'DERIVED_EMP_STATE',	'DERIVED_EMP_COUNTRY',	'DERIVED_EMP_GEOGRAPHY',	'ONS_OFF_FLAG',	'TOTAL_DAYS',	'AREA',	'ONSITE_DAYS',	'OFFSHORE_DAYS',	'BILLABLE',	'ONSITE_BILLABLE_DAYS',	'OFFSHORE_BILLABLE_DAYS',	'NON BILLABLE',	'ONSITE_NONBILLABLE_DAYS',	'OFFSHORE_NONBILLABLE_DAYS',	'SUPPORT',	'ONSITE_SERVICE_DAYS',	'OFFSHORE_SERVICE_DAYS',	'FREE',	'ONSITE_FREE_DAYS',	'OFFSHORE_FREE_DAYS',	'LEAVE_WHEN_BILLABLE_ONSITE',	'LEAVE_WHEN_BILLABLE_OFFSHORE',	'LEAVE_WHEN_NON_BILLABLE_ONSITE',	'LEAVE_WHEN_NON_BILLABLE_OFFSHORE',	'LEAVE_WHEN_FREE_ONSITE',	'LEAVE_WHEN_FREE_OFFSHORE',	'LEAVE_WHEN_SERVICE_ONSITE',	'LEAVE_WHEN_SERVICE_OFFSHORE',	'LEAVE_WHEN_OTHERS_ONSITE',	'LEAVE_WHEN_OTHERS_OFFSHORE',	'PRACTICE',	'PRAC_CC',	'PRAC_CC_DESC',	'SAP_PRAC_DESC',	'SAP_SUBPRAC_DESC',	'SEC_PRACTICE',	'SEC_PRAC_CC',	'SEC_PRAC_CC_DESC',	'SEC_SUPID',	'SUB_GEO',	'PREVIOUS_EXP_YRS',	'PREVIOUS_EXP_MON',	'WIPRO_EXP_YRS',	'WIPRO_EXP_MON',	'SOW_NO',	'LEAVE_STAT',	'INDICATOR',	'PERSONAL_AREA',	'PERSONAL_SUB_AREA',	'SUB_PRACTICE_DUAL_CREDIT',	'SUB_PRACTICE_DUAL_CREDIT_TEXT',	'CUSTOMER_ENGAGEMENT_TYPE',	'PROJECT_CLASSIFICATION',	'CAREER_BAND',	'START_DATE',	'END_DATE',	'CRMREFNO',	'PROJSTATUS',	'LOCATION',	'EMPLOYEE_EMAIL_ID',	'DATE_OF_JOINING',	'EXPERIENCE',	'BULGE',	'ROOKIE_STATUS',	'PRIMARY_SUPERVISOR',	'PROJECT_SUPERVISOR',	'CERTIFIED_SUITE',	'CERTIFIED_SKILL',	'POTENTIAL_SUITE',	'POTENTIAL_SKILL',	'UNASSESSED_SUITE',	'UNASSESSED_SKILL',	'PROJECT_ACQUIRED_SKILL',	'FLEXPROJ',	'NOCOST_LEAVEWHENBILLABLEONSITE',	'NOCOST_LEAVEWHENBILLABLEOFFSHORE',	'NOCOST_LEAVEWHENNONBILLABLEONSITE',	'NOCOST_LEAVEWHENNONBILLABLEOFFSHORE',	'NOCOST_LEAVEWHENFREEONSITE',	'NOCOST_LEAVEWHENFREEOFFSHORE',	'NOCOST_LEAVEWHENSERVICEONSITE',	'NOCOST_LEAVEWHENSERVICEOFFSHORE',	'PROJECT_SUBPRACTICE',	'PROJECT_SUBPRACTICE_DESC',	'HR_MRGID',	'PROJ_MRGID',	'ZZPROJCAT',	'DERIVED_SUITE_ID',	'DERIVED_SUITE_NAME',	'TOWER',	'INVESTMENT',	'INVESTED_FROM',	'INVESTED_TO',	'APPROVED_INVESTMENT',	'RTR_STATUS',	'RTR_DATE',	'INVESTMENT_FLAG',	'TRAINING_SKILLS',	'HOME_COST_CENTER',	'EMP_COST_CENTER_DESCRIPTION',	'GROUP_CUSTOMER_ID',	'GROUP_CUSTOMER_NAME',	'FRESHER_ENGAGEMENT_FLAG',	'BILLABLE_CATEGORY',	'ORDER_TYPE',	'ENTITY_TYPE',	'LARGE_PROGRAM',	'PRG_DIR_EMPID',	'PROGRAM',	'PGM_MGR_EMPID',	'PROJECT',	'PM_PL_EMPID',	'DM_EMPID',	'MODULE',	'MOD_EMPID',	'ROLE_DESCRIPTION',	'HOME_COMPANY_CODE',	'EMP_HOME_TOWN',	'RECENT_SKILLS',	'MSA_COUNTRY',	'MSA_COUNTRY_NAME',	'MSA_ROLE_ID',	'MSA_ROLE_DESCRIPTION',	'PROJECT_COMPANY_CODE',	'RESUME_PERCENTAGE',	'PORTFOLIO_ID',	'PORTFOLIO_NAME',	'DOMAIN_SUMMARY',	'LEGACY_PRAC_GROUP',	'LEGACY_COST_CENTER',	'LEGACY_PRACTICE',	'LEGACY_SUB_PRACTICE',	'LEGACY_SUB_PRAC_CODE',	'LEGACY_SUB_PRAC_TEXT',	'DU_PRACTICE',	'DU_SUB_PRACTICE',	'MODEL_TYPE',	'PROP_DATE',	'PROP_LVL',	'VIS_DATE',	'VIS_LVL',	'HOME_STRUCTURE',	'AGILE_FLAG',	'HOME_COUNTRY',	'HOME_GEOGRAPHY',	'PROJ_COUNTRY',	'PROJ_GEOGRAPHY',	'GEO_DU',	'FURLOUGH_WHEN_BILL_ONSITE',	'FURLOUGH_WHEN_BILL_OFFSHORE',	'FURLOUGH_WHEN_NONBILL_ONSITE',	'FURLOUGH_WHEN_NONBILL_OFFSHORE',	'FURLOUGH_WHEN_FREE_ONSITE',	'FURLOUGH_WHEN_FREE_OFFSHORE',	'FURLOUGH_WHEN_SERVICE_ONSITE',	'FURLOUGH_WHEN_SERVICE_OFFSHORE',	'DIGI_FLAG',	'FURLOUGH_FROM',	'FURLOUGH_TO',	'HR_EMP_ID',	'GIS_CATAPULT',	'SUPPORT_BILLABLE',	'COMPANY_IDENTIFICATION_FG',	'ONSITE_SUPPORT_CATEGORY',	'RM_EMP_NO',	'RM_EMP_NAME']]

Zcop1

#Rename column
Zcop4 = Zcop1.rename(columns={"LOAD_DATE": "DATE"})


#Data types of the data table.
Zcop1.dtypes
# Change Data type of column
Zcop1['Date'].astype(datetime64)

Zcop1 = Zcop1.astype(str)
Zcop1.to_feather("d:\\Zcop11.feather")

feather.write_dataframe(Zcop1, "d:\\Zcop4.feather")

writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')
Zcop2.to_excel(writer, sheet_name='Sheet1', index = False)
#df2.to_excel(writer,'Sheet2')
writer.save()
#Zcop2.info()
## Statistics of the table or column
#Zcop2.describe()

##-----------Data table size
#Zcop2.shape
## Selecting rows from data table
#Zcop2.iloc[0:5,:]
## Lable based row selection from data table
#Zcop1.loc[44,:]
## Sum the column value
[Zcop1['TOTAL_DAYS'].sum()]
## Logical based row selection from data table
Zcop1[Zcop1["EMP_CODE"] == "818088"]
## deleting columns
#Zcop5 = Zcop1.drop(columns="EMP_CODE")
## Dimension of data table
#Zcop2.ndim
## Selecting columns
Zcop2["EMP_NAME"]

Zcop2.to_pickle('Empname.pkl')
Zcop3 = pd.read_pickle('Empname.pkl') 

print(Zcop3)

#Zcop2.to_msgpack('foo.msg')

# Delete rows from the data table
#Zcop6 = Zcop1.set_index("EMP_CODE")
Zcop6 = Zcop1.drop("818088", axis=0)

# Delete Data table variable
del ERD1
del ERD4
del Zcop1['ERD_Date']