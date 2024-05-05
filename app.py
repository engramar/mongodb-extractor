from pymongo.mongo_client import MongoClient
from pymongo.server_api import ServerApi
from urllib.parse import quote_plus
from pymongo.mongo_client import MongoClient
from pymongo.server_api import ServerApi
import certifi
import pandas as pd

username = quote_plus('user')
password = quote_plus('pass')

cluster = 'test-cluster.xxxxx.mongodb.net'
authSource = 'admin'  # or appropriate authentication database
authMechanism = 'DEFAULT'
database_name = "test.user"

uri = f"mongodb+srv://{username}:{password}@{cluster}/{database_name}?authSource={authSource}&authMechanism={authMechanism}"

try:
    # Connect to MongoDB
    client = MongoClient(uri, tlsCAFile=certifi.where())
    
    db = client.get_database('caafg')
    print("Connected successfully to MongoDB")
    
    collection = db.get_collection('users')  # Assuming your collection name is 'users'
    print("Collection:", collection)
    print("Step 1: Collection retrieved")

    # Fetch all user documents from the collection
    alumni = collection.find()
    print("Alumni:", alumni)

    # Iterate over each alumni document and print its contents
    # for alumnus in alumni:
      #  print(alumnus)
    
    df = pd.DataFrame(alumni)
    # print(df)

    excel_file = "mongodb_toexcel.xlsx"
    df.to_excel(excel_file, index=False, engine='openpyxl')

except Exception as e:
    print("Error connecting to MongoDB:", e)    
