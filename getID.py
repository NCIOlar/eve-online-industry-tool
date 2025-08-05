import requests

def bulk_names_to_ids(names, language="zh"):
    url = "https://esi.evetech.net/latest/universe/ids/"
    headers = {
        "Accept-Language": language,
        "Content-Type": "application/json"
    }
    response = requests.post(url, json=names, headers=headers)
    response.raise_for_status()
    data = response.json()
    if "inventory_types" in data and len(data["inventory_types"]) > 0:
        return data["inventory_types"][0]["id"]
    else:
        return None


#Example usage
if __name__ == "__main__":
    names = ["皇冠蜥级蓝图"]
    result = bulk_names_to_ids(names)

    print(result)
