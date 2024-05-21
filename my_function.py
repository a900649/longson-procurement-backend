import json
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib, ssl
from PIL import Image
import pandas as pd
import variables as v
import pymysql
from datetime import datetime
import json
import os
import zipfile
from pymysql.converters import escape_string

def load_info_by_excel():

    title_df = pd.read_excel(v.excel_info_filename, engine="openpyxl", sheet_name="Title")
    introduction_df = pd.read_excel(v.excel_info_filename, engine="openpyxl", sheet_name="Requirements Introduction")
    common_col_df = pd.read_excel(v.excel_info_filename, engine="openpyxl", sheet_name="Common Columns")
    individual_col_df = pd.read_excel(v.excel_info_filename, engine="openpyxl", sheet_name="Individual Columns")
    function_df = pd.read_excel(v.excel_info_filename, engine="openpyxl", sheet_name="Function")
    verification_code_df = pd.read_excel(v.excel_info_filename, engine="openpyxl", sheet_name="Verification Code")
    internal_mail_receipients_df = pd.read_excel(v.excel_info_filename, engine="openpyxl", sheet_name="Internal Mail Receipients")
    supplier_mail_setting_df = pd.read_excel(v.excel_info_filename, engine="openpyxl", sheet_name="Supplier Mail Setting")
    photo_df = pd.read_excel(v.excel_info_filename, engine="openpyxl", sheet_name="Photo")
    check_df = pd.read_excel(v.excel_info_filename, engine="openpyxl", sheet_name="Check")

    # 修正Column Name 多出來的空格
    common_col_df["Column Name"] = common_col_df["Column Name"].apply(lambda x: x.strip())
    individual_col_df["Column Name"] = individual_col_df["Column Name"].apply(lambda x: x.strip())

    return title_df, introduction_df, common_col_df, individual_col_df, function_df, verification_code_df, internal_mail_receipients_df, supplier_mail_setting_df, photo_df, check_df

def load_info_by_sql(program_name):
    db = pymysql.connect(host=v.mysql_host, port=v.mysql_port, user=v.mysql_user, password=v.mysql_password, database=v.db_name)
    sql = "SELECT * FROM {}_info".format(program_name)
    cursor = db.cursor()
    k = cursor.execute(sql)
    data = cursor.fetchone()
    db.close()

    [title_df, introduction_df, common_col_df, individual_col_df, function_df, verification_code_df, internal_mail_receipients_df, supplier_mail_setting_df, photo_df, check_df] = \
        [pd.DataFrame(json.loads(i)["data"], columns=json.loads(i)["columns"], index=json.loads(i)["index"]) for i in data[:-1]]

    return title_df, introduction_df, common_col_df, individual_col_df, function_df, verification_code_df, internal_mail_receipients_df, supplier_mail_setting_df, photo_df, check_df

def upload_temp_row(name,data):

    db = pymysql.connect(host=v.mysql_host, port=v.mysql_port, user=v.mysql_user, password=v.mysql_password, database=v.db_name)
    cursor = db.cursor()
    sql = "REPLACE into {}_temp (Name, Data) values('{}', '{}')".format(v.program_name,name, escape_string(data))
    cursor.execute(sql)
    db.commit()
    db.close()

def load_temp_row(name):
    db = pymysql.connect(host=v.mysql_host, port=v.mysql_port, user=v.mysql_user, password=v.mysql_password, database=v.db_name)
    sql = "SELECT Data FROM {}_temp WHERE Name = '{}' ".format(v.program_name, name)
    cursor = db.cursor()
    k = cursor.execute(sql)
    data = cursor.fetchall()
    return data

def create_new_results_file(common_col_list, individual_col_list, product_list):
    col_list = [i[0] for i in v.first_col + common_col_list + individual_col_list + v.last_col]
    writer = pd.ExcelWriter(v.results_file_path)

    for product in product_list:
        results_df = pd.DataFrame([], columns=col_list)
        results_df.set_index("RowID", inplace=True, drop=True)
        results_df.to_excel(writer, sheet_name=product)

    writer.close()

def upload_results_table(sql_data_list_dict,columns_list):

    save_list = []
    for key in sql_data_list_dict.keys():
        sql_data_list = sql_data_list_dict[key]
        save_list.append(sql_data_list)
    print(save_list)
    if len(save_list) != 0:
        columns_str_list = ["`" + str(i) + "`" for i in columns_list]
        value_str_list = ["%s"] * len(save_list[0])
        db = pymysql.connect(host=v.mysql_host, port=v.mysql_port, user=v.mysql_user, password=v.mysql_password, database=v.db_name)
        cursor = db.cursor()
        sql = "REPLACE INTO {}_results ({}) VALUES ({})".format(v.program_name,",".join(columns_str_list), ",".join(value_str_list))
        k = cursor.executemany(sql, save_list)
        db.commit()
        db.close()

def load_results_table(program_name):

    title_df, introduction_df, common_col_df, individual_col_df, function_df, verification_code_df, internal_mail_receipients_df, supplier_mail_setting_df, photo_df, check_df = load_info_by_sql(program_name)

    columns_name_list = [i[0] for i in v.first_col] + list(common_col_df["Column Name"]) + list(individual_col_df["Column Name"]) + [i[0] for i in v.last_col]
    columns_name_str = ",".join(["`" + i + "`" for i in columns_name_list])

    db = pymysql.connect(host=v.mysql_host, port=v.mysql_port, user=v.mysql_user, password=v.mysql_password, database=v.db_name)
    sql = "SELECT {} FROM {}_results ".format(columns_name_str,program_name)
    cursor = db.cursor()
    k = cursor.execute(sql)
    data = cursor.fetchall()

    data_df = pd.DataFrame(data, columns=columns_name_list)
    data_df.set_index("RowID", inplace=True, drop=True)

    return data_df

def keep_new_data(data_df,col_list):
    if len(data_df) != 0:
        # 同一天相同店家多筆資料時，保留時間較晚的
        data_df.sort_values(by=["Update DateTime"], ascending=False, inplace=True)
        data_df.reset_index(inplace=True, drop=False)
        data_df = data_df.drop_duplicates(subset=col_list, keep="first")
        data_df.set_index("RowID",inplace=True,drop=True)
    return data_df

def attachment_file_zip(program_name):

    zf = zipfile.ZipFile('Attachment.zip', 'w', zipfile.ZIP_DEFLATED)
    for root, dirs, files in os.walk("Attachment//{}".format(program_name)):
        for file_name in files:
            zf.write(os.path.join(root, file_name))
    zf.close()