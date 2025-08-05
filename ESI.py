import requests
import math

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
        return [item["id"] for item in data["inventory_types"]]  # Return all IDs in a list
    else:
        return []

def get_names_from_id(item_id):
    url = "https://esi.evetech.net/latest/universe/names/"
    payload = [item_id]

    # Get English name
    headers_en = {
        "Accept-Language": "en",
        "Content-Type": "application/json"
    }
    response_en = requests.post(url, json=payload, headers=headers_en)
    response_en.raise_for_status()
    data_en = response_en.json()
    name_en = data_en[0]["name"] if data_en else ""

    # Get Chinese name
    headers_zh = {
        "Accept-Language": "zh",
        "Content-Type": "application/json"
    }
    response_zh = requests.post(url, json=payload, headers=headers_zh)
    response_zh.raise_for_status()
    data_zh = response_zh.json()
    name_zh = data_zh[0]["name"] if data_zh else ""

    return {"en": name_en, "zh": name_zh}


def get_jitasell(material_ids):
    url = "https://market.fuzzwork.co.uk/aggregates/"
    station_id = 60003760  # Jita 4-4

    # Filter out NaN and cast to int
    material_ids = [
        int(mid) for mid in material_ids
        if not (isinstance(mid, float) and math.isnan(mid))
    ]

    types = ",".join(map(str, material_ids))

    response = requests.get(f"{url}?station={station_id}&types={types}")
    data = response.json()

    # Extract min sell prices into a dictionary
    prices = {}
    for mat_id in material_ids:
        mat_id_str = str(mat_id)
        if mat_id_str in data and "sell" in data[mat_id_str]:
            prices[mat_id] = float(data[mat_id_str]["sell"].get("min", 0))
        else:
            prices[mat_id] = 0  # Default if no price found
    return prices


def get_adjusted_prices(item_ids):
    """
    Fetch adjusted prices from EVE ESI API and return a dict of {type_id: adjusted_price}.

    :param item_ids: List of item IDs to fetch prices for.
    :return: Dictionary {type_id: adjusted_price}
    """
    url = "https://esi.evetech.net/latest/markets/prices/"

    try:
        response = requests.get(url)
        response.raise_for_status()
        prices_data = response.json()

        # Filter and build dictionary
        adjusted_prices = {
            item["type_id"]: item.get("adjusted_price", 0)
            for item in prices_data
            if item["type_id"] in item_ids
        }

        return adjusted_prices

    except requests.exceptions.RequestException as e:
        print(f"Error fetching market prices: {e}")
        return {}


# Example usage
if __name__ == "__main__":
    print(get_adjusted_prices([37149,37153]))
