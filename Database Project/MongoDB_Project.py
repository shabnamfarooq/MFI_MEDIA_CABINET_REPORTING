from pymongo import MongoClient

# Connect to MongoDB (replace the URI with your MongoDB connection string)
client = MongoClient("mongodb://localhost:27017/")

# Select the database
db = client["smartphones_db"]

# Select the collections
brands_collection = db["brands"]
smartphones_collection = db["smartphones"]

# Count the number of documents in each collection
brands_count = brands_collection.count_documents({})
smartphones_count = smartphones_collection.count_documents({})

print(f"Count of rows in 'brands' collection: {brands_count}")
print(f"Count of rows in 'smartphones' collection: {smartphones_count}")

# Show a sample of 3 documents from each collection
brands_sample = brands_collection.find().limit(3)
smartphones_sample = smartphones_collection.find().limit(3)

print("\nSample documents from 'brands' collection:")
for doc in brands_sample:
    print(doc)

print("\nSample documents from 'smartphones' collection:")
for doc in smartphones_sample:
    print(doc)

# Use a join (aggregation pipeline) to combine data from both collections
joined_data = smartphones_collection.aggregate([
    {
        "$lookup": {
            "from": "brands",
            "localField": "BrandID",
            "foreignField": "BrandID",
            "as": "brand_info"
        }
    },
    {
        "$unwind": "$brand_info"
    },
    {
        "$project": {
            "SmartphoneID": 1,
            "Smartphone": 1,
            "Model": 1,
            "RAM": 1,
            "Storage": 1,
            "Color": 1,
            "Final_Price": 1,
            "BrandName": "$brand_info.BrandName",
            "_id": 0
        }
    },
    {
        "$limit": 3
    }
])

print("\nJoined data (sample of 3 documents):")
for doc in joined_data:
    print(doc)
