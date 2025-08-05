import yaml
import pandas as pd
import getID
import ESI
import time
start_time = time.time()
#Load YAML files
# items = yaml.safe_load(open("types.yaml", "r", encoding="utf-8"))
blueprints = yaml.safe_load(open("blueprints.yaml", "r", encoding="utf-8"))

def generate_blueprint_excel(item_ID):

    bp = blueprints.get(item_ID)
    if not bp:
        print(f"No blueprint found for {item_ID}")
        return

    materials = bp.get("activities", {}).get("manufacturing", {}).get("materials", [])
    result = []
    for mat in materials:
        mat_id = mat["typeID"]
        quantity = mat["quantity"]
        material_info = ESI.get_names_from_id(mat_id)
        bp_info = ESI.get_names_from_id(item_ID)

        result.append({
            "Blueprint (ZH)": bp_info["zh"],
            "Blueprint (EN)": bp_info["en"],
            "Material (ZH)": material_info["zh"],
            "Material (EN)": material_info["en"],
            "Quantity": quantity
        })

    if not result:
        print(f"No blueprint found for {item_ID}")
        return

    df = pd.DataFrame(result)
    file_name = f"{item_ID}_materials.xlsx"
    df.to_excel(file_name, index=False)
    print(f"✅ Excel generated: {file_name}")




# # Example usage
names = ["伊什塔级蓝图"]
item_id = ESI.bulk_names_to_ids(names)
generate_blueprint_excel(item_id)
# if __name__ == "__main__":
#     names = ["皇冠蜥级蓝图"]
#     item_id = getID.bulk_names_to_ids(names)
#     print(item_id)

end_time = time.time()  # End timer
print(f"Execution time: {end_time - start_time:.2f} seconds")
