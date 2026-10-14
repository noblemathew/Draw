
users = [
    {"id": 101, "name": "Arun Mathew", "role": "Data Analyst", "status": "Active"},
    {"id": 102, "name": "Amy John", "role": "Software Developer", "status": "Active"},
    {"id": 103, "name": "Rahul Thomas", "role": "Python Developer", "status": "Inactive"},
    {"id": 104, "name": "Neha George", "role": "Data Engineer", "status": "Active"},
]

def search_users(keyword):
    results = [
        user for user in users
        if keyword.lower() in user["name"].lower()
        or keyword.lower() in user["role"].lower()
    ]

    if results:
        for user in results:
            print(
                f"ID: {user['id']} | "
                f"Name: {user['name']} | "
                f"Role: {user['role']} | "
                f"Status: {user['status']}"
            )
    else:
        print("No matching users found.")

search_users("Python")
