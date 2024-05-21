
# 資料存放位置
# data_source = "Excel"
data_source = "Azure SQL"

# 專案名稱
program_name = "acetic_acid_2nd_half_2024"

# 存在可下載的專案名稱
program_name_list = ["acetic_acid_2nd_half_2024"]

# 網頁Title
page_title = "Longson Procurement System"

system_photo_path = "System Photo/{}"

logo_filename = 'Longson.jpg'
icon_filename = 'Smile.webp'
form_tail_filename = 'Thanks.jpg'

# 網頁Title
page_title = "Longson Procurement System"

# Excel的Info檔案名稱
excel_info_filename = "Longson Purchase Requirements Info.xlsx"

# Excel的結果檔案放置位置
results_file_path = "Procurement Results {}.xlsx".format(program_name)

# 其他資料存放區
temp_path = "Temp" + "/" + program_name
temp_data_path = temp_path + "//{}.json"
attachment_path = "Attachment" + "/" + program_name + "/{}" + "/{}"
system_photo_path = "System Photo/{}"


# DB 資訊
mysql_host = "longson.mysql.database.azure.com"
mysql_user='paul'
mysql_password='Yunxuan123'
mysql_port = 3306
db_name = "longson_procurement"

# Results Other Columns
first_col = [["RowID","Text"],["Product","Text"]]
last_col = [["Attachment","Text"],["Update DateTime","DateTime"],["Verification Code","Text"],["Verification Code Name","Text"]]

# BLOB
blob_connection_string = "DefaultEndpointsProtocol=https;AccountName=kenso;AccountKey=Wto5Ig361Z/aVQuxEfvM7b9MnKi3IctRB70fq5X53CCLlQ84BpFaS9T5HWVLcFwOVSEcljz0Aa40+AStQgHifw==;EndpointSuffix=core.windows.net"
blob_container = "longson"