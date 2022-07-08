import os
import pandas as pd
from functools import reduce

import smtplib,ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

# Calculate the path to the root directory of this script
ROOT_DIR = os.path.realpath(os.path.join(os.path.dirname(__file__), '.'))

# Import SNAP-Ed staff

FY22_INEP_Staff = pd.ExcelFile(ROOT_DIR + "/sample_inputs/FY22_INEP_Staff_List.xlsx")
# Alternatively, use the absolute path to the staff list
# FY22_INEP_Staff = pd.ExcelFile(r"C:\Users\netid\Box\INEP Staff Lists\FY22 INEP Staff List.xlsx")
# Adjust header argument below for actual staff list
SNAP_Ed_Staff = pd.read_excel(FY22_INEP_Staff, sheet_name='SNAP-Ed Staff List', header=0) 
SNAP_Ed_Staff['NAME'] = SNAP_Ed_Staff['NAME'].str.strip()
SNAP_Ed_Staff['E-MAIL'] = SNAP_Ed_Staff['E-MAIL'].str.strip()

# Import CPHP staff

# Adjust header argument below for actual staff list
CPHP_Staff = pd.read_excel(FY22_INEP_Staff, sheet_name='CPHP Staff List', header=0).rename(columns={'Last Name' : 'last_name',
                                                                                                    'First Name' : 'first_name',
                                                                                                    'email Address' : 'email'})
CPHP_Staff['full_name'] = CPHP_Staff['first_name'].map(str) + ' ' + CPHP_Staff['last_name'].map(str)
CPHP_Staff = CPHP_Staff.loc[CPHP_Staff['email'].notnull(), ['full_name', 'email']]
CPHP_Staff['email'] = CPHP_Staff['email'].str.strip()

# Import PEARS users

PEARS_User_Export = pd.read_excel(ROOT_DIR + "/sample_inputs/User_Export.xlsx", sheet_name='User Data')
# Alternatively, use the absolute path to the user export
# PEARS_User_Export = pd.read_excel(r"\path\to\User_Export.xlsx", sheet_name='User Data')
PEARS_User_Export = PEARS_User_Export.loc[PEARS_User_Export['is_active'] == 1]

# Function for merging PEARS module records with collaborator data
# df: dataframe of module data
# module_id: string for the module's id column label
# excel_file: pandas.ExcelFile of the module export
def merge_collaborators(df, module_id, excel_file): 
    collaborators = pd.read_excel(excel_file, 'Collaborators')
    collaborators = pd.merge(collaborators, PEARS_User_Export, how='left', left_on='user', right_on='full_name')
    collaborators = pd.merge(collaborators, df, how='left', on=module_id)
    collaborators = collaborators.loc[:, [module_id, 'user', 'email', 'created', 'modified']]
    return collaborators

# Desired modules to report on
# 'Excel_File', 'Sheet Name'

import_modules = [['Program_Activities', 'Program Activity Data'],
                  ['Indirect_Activity', 'Indirect Activity Data'],
                  ['Coalition', 'Coalition Data'],
                  ['Partnership', 'Partnership Data'],
                  ['PSE_Site_Activity', 'PSE Data'],
                  ['Success_Story', 'Success Story Data']]

# Id column labels for each module in import_modules

module_ids = ['program_id', 'activity_id', 'coalition_id', 'partnership_id', 'pse_id', 'story_id']

# Import record creation and collaboration data for each module

module_dfs = []

for index, item in enumerate(import_modules):   
    wb = pd.ExcelFile(ROOT_DIR + "/sample_inputs/" + item[0] + "_Export.xlsx")
    # Record creation data
    # Module records aggregated by the user specified in the 'reported_by' field
    create_df = pd.read_excel(wb, item[1])
    # Colloboration data
    # Module records aggregated by the user(s) specified in the 'collaborators' field
    collab_df = merge_collaborators(create_df, module_ids[index], wb)
    module_dfs.append([create_df, collab_df])

  
# Create PEARS SNAP-Ed Staff Report

# Null values in FY22 INEP Staff List.xlsx 
staff_nulls = ('N/A', 'NEW', 'OPEN')
# Prep dataframe of SNAP-Ed staff
staff = SNAP_Ed_Staff.loc[~SNAP_Ed_Staff['NAME'].isin(staff_nulls) & SNAP_Ed_Staff['NAME'].notnull(), ['UNIT #', 'JOB CLASS', 'NAME', 'E-MAIL']]
staff = staff.loc[staff['E-MAIL'].notnull()]
staff = staff.rename(columns={'E-MAIL' : 'email'})

# Timestamp for day the report is run
ts = pd.to_datetime("today").date()
# PeriodArray/Index object for report month
prev_month = (ts - pd.DateOffset(months=1)).to_period('M')
# Start date of the report period
prev_month_lb = (ts.replace(day=1) - pd.DateOffset(months=1)).date()
# End date of the report period
# Prior month's records are typically entered by the 10th day of subsequent month
prev_month_ub = ts.replace(day=10)

# Function to create list of dataframes consisting of
# counts of record creation data created during the previous month,
# counts of record creation data modified during the previous month,
# counts of record creation data created during the current year to date,
# counts of collaboration data created during the previous month,
# counts of collaboration data modified during the previous month,
# counts of collaboration data created during the current year to date
# df_created: dataframe of module record creation data
# df_collab: dataframe of module collaboration data
# module_id: string for the module's id column label
# prev_month_lb: datetime.date object for the start date of the report period
# prev_month_ub: datetime.date object for the end date of the report period
def created_collab_dfs(df_created, df_collab, module_id, prev_month_lb=prev_month_lb, date_ub=prev_month_ub):
	df_created = df_created.rename(columns={'reported_by_email' : 'email'})
	df_created['created'] = pd.to_datetime(df_created['created']).dt.date
	Prev_MO_Created = df_created.loc[(df_created['created'] >= prev_month_lb) & (df_created['created'] <= prev_month_ub)].groupby('email')[module_id].count().reset_index(name='Prev_MO_Created')

	df_created['modified'] = pd.to_datetime(df_created['modified']).dt.date
	Prev_MO_Modified = df_created.loc[(df_created['modified'] >= prev_month_lb) & (df_created['modified'] <= prev_month_ub)].groupby('email')[module_id].count().reset_index(name='Prev_MO_Modified')

	YTD_Created = df_created.groupby('email')[module_id].count().reset_index(name='YTD_Created')
		
	df_collab['created'] = pd.to_datetime(df_collab['created']).dt.date
	Prev_MO_Collab = df_collab.loc[(df_collab['created'] >= prev_month_lb) & (df_collab['created'] <= prev_month_ub)].groupby('email')[module_id].count().reset_index(name='Prev_MO_Collab')

	df_collab['modified'] = pd.to_datetime(df_collab['modified']).dt.date
	Prev_MO_Collab_Mod = df_collab.loc[(df_collab['modified'] >= prev_month_lb) & (df_collab['modified'] <= prev_month_ub)].groupby('email')[module_id].count().reset_index(name='Prev_MO_Collab_Mod')

	YTD_Collab = df_collab.groupby('email')[module_id].count().reset_index(name='YTD_Collab')

	dfs = [Prev_MO_Created, Prev_MO_Modified, YTD_Created, Prev_MO_Collab, Prev_MO_Collab_Mod, YTD_Collab]
	return dfs

# Desired modules to report on

modules = ['Program Activities', 'Indirect Activities', 'Coalitions', 'Partnerships', 'PSE', 'Success Stories']

# For each module, aggregate record creation/collaboration counts by each timeframe 

module_created_collab_dfs = []

for index, item in enumerate(module_dfs):   
    module_created_collab_dfs.append(created_collab_dfs(item[0], item[1], module_ids[index]))

# Function to merge record counts and staff list
# dfs: list of dataframes returned from created_collab_dfs()
# staff: dataframe of staff
# module: string of the module name
def module_staff_entries(dfs, staff, module):
	dfs = [staff] + dfs

	df_merged = reduce(lambda left, right: pd.merge(left, right, how='left', on='email'), dfs)
	df_merged = df_merged.fillna(0)

	df_merged = df_merged.rename(columns={'Prev_MO_Created' : module + ' Created ' + prev_month.strftime('%b-%Y'),
										  'Prev_MO_Modified' : module + ' Modified ' + prev_month.strftime('%b-%Y'),
										  'YTD_Created' : module + ' Created YTD',
										  'Prev_MO_Collab' : module + ' Collaborated Created ' + prev_month.strftime('%b-%Y'),
										  'Prev_MO_Collab_Mod' : module + ' Collaborated Modified ' + prev_month.strftime('%b-%Y'),
										  'YTD_Collab' : module + ' Collaborated Created YTD'})
	return df_merged

# Merge record counts for each module with SNAP-Ed staff

extension_staff_modules = []

for index, item in enumerate(module_created_collab_dfs):   
    extension_staff_modules.append(module_staff_entries(item, staff, modules[index]))


# Function to compile the staff report formatted to each agency's specifications
# dfs: list of record count dfs returned from module_staff_entries()
# agency: string, either 'Extension' or 'CPHP'
def compile_report(dfs, agency='Extension'):
    if agency == 'Extension':
        staff_cols = ['UNIT #','JOB CLASS', 'NAME', 'email']
        sort_cols = ['UNIT #', 'NAME']
        rename_cols = {'UNIT #' : 'Unit #', 'JOB CLASS' : 'Job Class', 'NAME' : 'Name', 'Email' : 'email'} 
    elif agency == 'CPHP':
        staff_cols = ['full_name','email']
        sort_cols = ['full_name']
        rename_cols = {'full_name' : 'Name', 'email' : 'Email'}
        
    report = reduce(lambda left, right: pd.merge(left, right, how='outer', on=staff_cols), dfs)
    
    report = report.sort_values(by=sort_cols )
    report['Total Entries Created ' + prev_month.strftime('%b-%Y')] = report.loc[:, report.columns.str.contains('Created ' + prev_month.strftime('%b-%Y')) & ~report.columns.str.contains('Collaborated ')].sum(axis=1)
    report['Total Entries Modified ' + prev_month.strftime('%b-%Y')] = report.loc[:, report.columns.str.contains('Modified ' + prev_month.strftime('%b-%Y')) & ~report.columns.str.contains('Collaborated ')].sum(axis=1)
    report['Total Entries Created YTD'] = report.loc[:, report.columns.str.contains('Created YTD') & ~report.columns.str.contains('Collaborated ')].sum(axis=1)
    report['Total Entries Collaborated Created ' + prev_month.strftime('%b-%Y')] = report.loc[:, report.columns.str.contains('Collaborated Created ' + prev_month.strftime('%b-%Y'))].sum(axis=1)
    report['Total Entries Collaborated Modified ' + prev_month.strftime('%b-%Y')] = report.loc[:, report.columns.str.contains('Collaborated Modified ' + prev_month.strftime('%b-%Y'))].sum(axis=1)
    report['Total Entries Collaborated Created YTD'] = report.loc[:, report.columns.str.contains('Collaborated Created YTD')].sum(axis=1)
    # Set boolean column for staff who have 0 entries for the month
    report['0 Entries'] = False
    total_prev_month_columns = report.columns[report.columns.str.contains('Total Entries') & ~report.columns.str.contains('YTD')]
    report.loc[(report.filter(items=total_prev_month_columns) == 0).all(1), '0 Entries'] = True
    
    zero_entries_index = len(staff_cols) - 2
    report.insert(zero_entries_index, '0 Entries', report.pop('0 Entries'))
    report = report.rename(columns=rename_cols)
    
    staff_cols_index = len(staff_cols) + 1
    cols = report.columns.tolist()
    cols = cols[:staff_cols_index] + cols[-6:] + cols[staff_cols_index:-6]
    report = report[cols]
    return report

# Compiled staff report for Extension (SNAP-Ed)
extension_report = compile_report(extension_staff_modules)

# Function to export the staff report as an xlsx formatted to each agency's specifications
# dfs: dict of sheet name and dataframe returned from compile_report()
# file_path: string for the output directory and filename
# agency: string, either 'Extension' or 'CPHP'
def save_staff_report(dfs, file_path, agency='Extension'):
    if agency == 'Extension':
        freeze_cols = 5
        cond_form = [3, '=C1=TRUE']
    elif agency == 'CPHP':
        freeze_cols = 3
        cond_form = [1, '=A1=TRUE']
        
    writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
    for sheetname, df in dfs.items():
        df.to_excel(writer, sheet_name=sheetname, index=False, freeze_panes=(1, freeze_cols))
        worksheet = writer.sheets[sheetname]
        workbook  = writer.book
        red_format = workbook.add_format({'bg_color' : '#FFC7CE', 'font_color' : '#9C0006'})
        worksheet.autofilter(0, 0, 0, len(df.columns)-1)
        # Highlight staff who have 0 entries for the month
        worksheet.conditional_format(0, cond_form[0], len(df), cond_form[0],
                               {'type':     'formula',
                                'criteria': cond_form[1],
                                'format':   red_format})
        
        for idx, col in enumerate(df):
            series = df[col]
            max_len = max((
				series.astype(str).map(len).max(),  # len of largest item
				len(str(series.name))  # len of column name/header
				)) + 1
            worksheet.set_column(idx, idx, max_len)
        
        writer.close()

# Save extension report

dfs = {'Extension Staff PEARS Entries' : extension_report}
filename1 = 'Extension Staff PEARS Entries ' + prev_month.strftime('%Y-%m') + '.xlsx'
out_path = ROOT_DIR + "/sample_outputs"
# Alternatively, use the absolute path to the output directory
# out_path = r"C:\Users\netid\path\to\output\directory"
file_path1 = out_path + '/' + filename1

save_staff_report(dfs, file_path1, agency='Extension')


# Create PEARS CPHP Staff Report


cphp_staff_modules = []

for index, item in enumerate(module_created_collab_dfs):   
    cphp_staff_modules.append(module_staff_entries(item, CPHP_Staff, modules[index]))


cphp_report = compile_report(cphp_staff_modules, agency='CPHP')

dfs = {'CPHP Staff PEARS Entries' : cphp_report}
filename2 = 'CPHP Staff PEARS Entries ' + prev_month.strftime('%Y-%m') + '.xlsx'
file_path2 = out_path + '/' + filename2

save_staff_report(dfs, file_path2, agency='CPHP')


# Email Reports

# Set the following variables with the appropriate credentials and recipients
username='your_username@domain.com'
password='your_password'
send_from = 'your_username@domain.com'
Cc = 'list@domain.com, of_recipients@domain.com'

# Function to send an email with an xlsx attachment
# send_from: string for the sender's email address
# send_to: string for the recipient's email address
# Cc: string of comma-separated cc addresses
# subject: string for the email subject line
# text: string for the email body
# file_path: string for the xlsx attachment's filepath
# filename: string for the xlsx attachments filename
# username: string for the username to athenticate with
# password: string for the password to authenticate with
# isTls: boolean, True to put the SMTP connection in Transport Layer Security mode
def send_mail(send_from, send_to, Cc, subject, text, file_path, filename, username, password, isTls=True):
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = send_to
    msg['Cc'] = Cc
    msg['Date'] = formatdate(localtime = True)
    msg['Subject'] = subject
    msg.attach(MIMEText(text))

    fp = open(file_path, 'rb')

    part = MIMEBase('application','vnd.ms-excel')
    part.set_payload(fp.read())
    fp.close()
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment', filename=filename)
    msg.attach(part)

    smtp = smtplib.SMTP('smtp.office365.com', 587)
    if isTls:
        smtp.starttls()
    try:    
        smtp.login(username,password)
        smtp.sendmail(send_from, send_to.split(',') + msg['Cc'].split(','), msg.as_string())
    except smtplib.SMTPAuthenticationError:
        print("Authentication failed. Make sure to provide a valid username and password.")
    smtp.quit()

# Email the SNAP-Ed staff report

send_to1 = 'snap_ed_recipient@domain.com'    
subject1 = 'Extension Staff PEARS Entries ' + prev_month.strftime('%Y-%m')
text1 = subject1 + ' attached.'

send_mail(send_from, send_to1, Cc, subject1, text1, file_path1, filename1, username, password, isTls=True)

# Email the CPHP staff report

send_to2 = 'cphp_recipient@domain.com' 
subject2 = 'CPHP Staff PEARS Entries ' + prev_month.strftime('%Y-%m')
text2 = subject2 + ' attached.'

send_mail(send_from, send_to2, Cc, subject2, text2, file_path2, filename2, username, password, isTls=True)
