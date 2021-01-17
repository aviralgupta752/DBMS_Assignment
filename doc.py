import mysql.connector as mysql
from docx.api import Document
import pprint
table_names = ["Item_Table", "Customer_Table", "Salesman_Table", "Sales_Order_Table", "Customer_Table1", "Loan_Table", "Borrower_Table", "Account_Table", "Depositor_Table"]
document = Document("/home/avi/Downloads/file.docx")

# CREATING A DATABSE NAMED STORE
database_name = "store"
db = mysql.connect(
	host = "localhost", 
	user = "root", 
	passwd = "your_password",
	)
cursor = db.cursor()
try:
	cursor.execute("CREATE DATABASE " + database_name)
	cursor.execute("USE " + database_name)
except:
	pass
db.commit()
###############################################################

# USING THIS LITTLE CODE TO DELETE ALL THE TABLES, IN CASE WE WANT TO RE-RUN THE CODE ON THE SAME DATABSE WITH THE SAME TABLES
for i in range(len(document.tables)):
	db = mysql.connect(
		host = "localhost", 
		user = "root", 
		passwd = "your_password",
		database = database_name
		)
	cursor = db.cursor()
	try:
		cursor.execute("DROP TABLE " + table_names[i]);
	except:
		pass
	db.commit();
###############################################################

for i in range(len(document.tables)):
	# MAKING A LIST OF ALL THE DATA IN A TABLE
	table = document.tables[i]
	data = []
	keys = None
	for j, row in enumerate(table.rows):
		text = (cell.text for cell in row.cells)
		if(j==0):
			keys = tuple(text)
			continue
		row_data = tuple(text)
		data.append(row_data)

	# pprint.pprint(data)
	###############################################################

	# CREATING AND INSERTING DATA INTO A TABLE
	db = mysql.connect(
		host = "localhost", 
		user = "root", 
		passwd = "your_password",
		database = database_name
		)

	cursor = db.cursor()
	create_query = "CREATE TABLE " + table_names[i] + " ("

	for j in range(len(keys)):
		item = str(keys[j]).strip().replace(" ","").replace("\n","") + " VARCHAR(30) NOT NULL, "
		create_query += item
		if(j == len(keys)-1):
			create_query = create_query[:len(create_query)-2] + ");"

	# print(create_query)
	cursor.execute(create_query)

	insert_query = "INSERT INTO " + table_names[i] + " VALUES ("
	for j in range(len(data)):
		row_data = data[j]
		item = ""
		for k in row_data:
			item += "'" + str(k) + "', "
		item = item[:-2] + ');'
		# print(insert_query + item)
		cursor.execute(insert_query + item)

	db.commit()
	###############################################################
	print("Table " + str(i) + " complete.")

