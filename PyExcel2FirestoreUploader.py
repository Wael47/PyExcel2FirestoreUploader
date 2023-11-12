import pandas as pd
import firebase_admin
from firebase_admin import credentials
from firebase_admin import firestore

cred = credentials.Certificate('rehlate-application-firebase-adminsdk-n6i8h-af530cd26c.json')

app = firebase_admin.initialize_app(cred)

db = firestore.client()

cities_ref = db.collection("tour")

df = pd.read_excel('Dataset.xlsx')

start_row = 0
end_row = 44
start_col = 0
end_col = 11
df = df.iloc[start_row:end_row + 1, start_col:end_col + 1]
headers = df.columns

for index, row in df.iterrows():
    json = ""
    json += '{'
    row = row[::-1]
    for column_name, cell_value in row.items():
        column_name, cell_value = str(column_name).strip(), str(cell_value).strip()
        if column_name == 'الرقم' or column_name == 'التقييم':
            json += f'"{column_name}":{int(float(cell_value))},'
        elif column_name == 'النشاطات' or column_name == 'مسار الرحلة':
            json += f'"{column_name}":{[j.strip() for j in cell_value.split(",") if j]},'
        elif column_name == 'التاريخ':
            json += f'"{column_name}":"{cell_value.split( )[0]}",'
        else:
            json += f'"{column_name}":"{cell_value}",'
    json += '}'
    final_dictionary = eval(json)
    print(final_dictionary)
    cities_ref.document().set(final_dictionary)

