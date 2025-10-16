import math

import numpy as np
import yaml
import pandas as pd
import getID
import ESI
import time
start_time = time.time()

#open excel
items = pd.read_excel("eve_items.xlsx", sheet_name="item id")
blueprints = pd.read_excel("eve_items.xlsx", sheet_name="blueprint")
items_to_path = pd.read_excel("eve_items.xlsx", sheet_name="item to path")
reactions = pd.read_excel("eve_items.xlsx", sheet_name="reaction formula")
reaction_to_path = pd.read_excel("eve_items.xlsx", sheet_name="reaction to path")
# Set "id" column as index for faster lookup
# items.set_index("id", inplace=True)
blueprints.set_index("Blueprint ID", inplace=True)
reactions.set_index("Item ID", inplace=True)


def generate_blueprint_sheet(prev_sheet, count, price_control):

    #get material from previous sheet
    material_ids = prev_sheet["Material ID"].tolist()
    # Step 2: Filter items_to_path by matching material IDs to Item ID
    matched_items = items_to_path[items_to_path["Item ID"].isin(material_ids)]
    # Step 3: Get corresponding Item Path IDs
    item_path_ids = matched_items["Item Path ID"].unique().tolist()
    if len(item_path_ids) == 0:
        return [], 1
    # Step 4: Filter blueprints by these blueprint IDs
    component_df = blueprints.loc[item_path_ids]

    component_df["材料效率"] = 1

    #component_df["Base Name"] = component_df["Blueprint EN"].str.replace(" Blueprint", "", regex=False)
    # Create a mapping dictionary from column E to column B in sheet2
    mapping = dict(zip(items_to_path["Item Path EN"], items_to_path["Item EN"]))
    # Map column A in sheet1 to the corresponding product EN value in sheet2
    component_df["Base Name"] = component_df["Blueprint EN"].map(mapping)
    # Create a mapping from Material EN to 总需求
    mapping = dict(zip(prev_sheet["Unique Material EN"], prev_sheet["Rounded 总需求"]))
    component_df["总需求"] = component_df["Base Name"].map(mapping)

    component_df = component_df.reset_index(drop=True)
    for i, (idx, row) in enumerate(component_df.iterrows(), start=2):  # start=2 for Excel rows
        existing_formula = str(row["总需求"])  # convert to string to avoid issues
        component_df.at[idx, "总需求"] = f"={existing_formula[1:]}*ROUNDUP(MaterialBlueprints{count}!F{i}*MaterialBlueprints{count}!G{i},0)"
    #Drop base name helper colume
    component_df = component_df.drop(columns=["Base Name"])


    component_df[""] = None
    component_df[" "] = None
    # Group by Material ID but keep the first ZH and EN names
    grouped_df = (
        component_df.groupby("Material ID", as_index=False)
        .agg({
            "总需求": "sum",
            "Material ZH": "first",
            "Material EN": "first"
        })
    )
    # Apply the fix_formulas function to "总需求"
    grouped_df["总需求"] = grouped_df["总需求"].apply(fix_formulas)

    # Add the grouped data back to the original sheet as new columns
    component_df["Unique Material ID"] = grouped_df["Material ID"]
    component_df["Unique Material ZH"] = grouped_df["Material ZH"]
    component_df["Unique Material EN"] = grouped_df["Material EN"]
    component_df["Summed 总需求"] = grouped_df["总需求"]

    def round_up_to_multiple(item_id, total):
        output = output_map.get(item_id)

        # If no output multiple, return as is
        if not output:
            return total

        # If it's an Excel formula string, inject ROUNDUP logic
        if isinstance(total, str) and total.strip().startswith("="):
            return f"=ROUNDUP(({total[1:]})/{output}, 0)*{output}"

    output_map = dict(zip(reaction_to_path["product ID"], reaction_to_path["Product Output"]))
    component_df["Rounded 总需求"] = component_df.apply(
        lambda row: round_up_to_multiple(row["Unique Material ID"], row["Summed 总需求"]), axis=1
    )

    component_df["  "] = None
    component_df.reset_index()
    material_ids = component_df["Unique Material ID"].tolist()
    min_sell_prices = ESI.get_jitasell(material_ids)
    # Assign prices
    component_df["Min Jita Sell"] = component_df["Unique Material ID"].map(min_sell_prices)

    component_df = component_df.reset_index(drop=True)
    for i, (idx, row) in enumerate(component_df.iterrows(), start=2):
        if pd.isna(row["Rounded 总需求"]) or str(row["Rounded 总需求"]).strip() == "":
            continue  # skip this row

        existing_formula = str(row["Rounded 总需求"])
        component_df.at[idx, "Total Cost"] = f"=({existing_formula[1:]})*(MaterialBlueprints{count}!Q{i})*({price_control})"

    component_df.at[0, "Total Cost Sum"] = f"=SUM(R{2}:R{len(component_df) + 1})"

    # Filter out rows where Total Cost is not empty or zero
    valid_rows = component_df[component_df["Total Cost"].notna() & (component_df["Total Cost"] != 0)]

    # Initialize 已完成数量 column
    component_df["已完成数量"] = ""
    start_row = 2
    # Fill values only for valid rows
    for i, idx in enumerate(valid_rows.index, start=start_row):
        component_df.at[idx, "剩余数量"] = f'=O{i}-T{i}'
        component_df.at[idx, "已完成数量"] = 0

    #get EVI
    component_df["   "] = None
    component_df["    "] = None
    component_df["     "] = None
    component_df.reset_index()
    material_ids = component_df["Unique Material ID"].tolist()
    min_sell_prices = ESI.get_adjusted_prices(material_ids)
    # Assign prices
    component_df["Adjusted Price"] = component_df["Unique Material ID"].map(min_sell_prices)

    component_df = component_df.reset_index(drop=True)
    for i, (idx, row) in enumerate(component_df.iterrows(), start=2):
        if pd.isna(row["Rounded 总需求"]) or str(row["Rounded 总需求"]).strip() == "":
            continue  # skip this row

        existing_formula = str(row["Rounded 总需求"])
        component_df.at[idx, "Total Adjusted Cost"] = f"=({existing_formula[1:]})*MaterialBlueprints{count}!Y{i}"
    component_df.at[0, "Total Adjusted Cost Sum"] = f"=SUM(MaterialBlueprints{count}!Z{2}:MaterialBlueprints{count}!Z{len(component_df) + 1})"

    return component_df, 0



def fix_formulas(formula):
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula
    parts = formula.split("=")
    return "=" + "+".join(filter(None, parts[1:]))  # Keep the first =, turn others into +


def generate_reaction_sheet(prev_sheet,count, price_control):

    #get material from previous sheet
    material_ids = prev_sheet["Material ID"].tolist()
    # Step 2: Filter items_to_path by matching material IDs to Item ID
    matched_items = reaction_to_path[reaction_to_path["product ID"].isin(material_ids)]
    # Step 3: Get corresponding Item Path IDs
    item_path_ids = matched_items["Item Path ID"].unique().tolist()
    if len(item_path_ids) == 0:
        return [], 1
    # Step 4: Filter blueprints by these blueprint IDs
    reaction_df = reactions.loc[item_path_ids]

    reaction_df["材料效率"] = 1

    # Create a mapping dictionary from column E to column B in sheet2
    mapping = dict(zip(reaction_to_path["Reaction Formula EN"], reaction_to_path["product EN"]))
    # Map column A in sheet1 to the corresponding product EN value in sheet2
    reaction_df["Base Name"] = reaction_df["Reaction Formula EN"].map(mapping)
    # Create a mapping from Material EN to 总需求
    mapping = dict(zip(prev_sheet["Unique Material EN"], prev_sheet["Rounded 总需求"]))
    reaction_df["总需求"] = reaction_df["Base Name"].map(mapping)

    reaction_df = reaction_df.reset_index(drop=True)
    for i, (idx, row) in enumerate(reaction_df.iterrows(), start=2):  # start=2 for Excel rows
        existing_formula = str(row["总需求"])  # convert to string to avoid issues
        reaction_df.at[idx, "总需求"] = f"=(({existing_formula[1:]})/ReactionFormulas{count}!C{i})*ROUNDUP(ReactionFormulas{count}!G{i}*ReactionFormulas{count}!H{i},0)"
    #Drop base name helper colume
    reaction_df = reaction_df.drop(columns=["Base Name"])


    reaction_df[""] = None
    reaction_df[" "] = None
    # Group by Material ID but keep the first ZH and EN names
    grouped_df = (
        reaction_df.groupby("Material ID", as_index=False)
        .agg({
            "总需求": "sum",
            "Material ZH": "first",
            "Material EN": "first"
        })
    )
    # Apply the fix_formulas function to "总需求"
    grouped_df["总需求"] = grouped_df["总需求"].apply(fix_formulas)

    # Add the grouped data back to the original sheet as new columns
    reaction_df["Unique Material ID"] = grouped_df["Material ID"]
    reaction_df["Unique Material ZH"] = grouped_df["Material ZH"]
    reaction_df["Unique Material EN"] = grouped_df["Material EN"]
    reaction_df["Summed 总需求"] = grouped_df["总需求"]

    def round_up_to_multiple(item_id, total):
        output = output_map.get(item_id)

        # If no output multiple, return as is
        if not output:
            return total

        # If it's an Excel formula string, inject ROUNDUP logic
        if isinstance(total, str) and total.strip().startswith("="):
            return f"=ROUNDUP(({total[1:]})/{output}, 0)*{output}"

    output_map = dict(zip(reaction_to_path["product ID"], reaction_to_path["Product Output"]))
    reaction_df["Rounded 总需求"] = reaction_df.apply(
        lambda row: round_up_to_multiple(row["Unique Material ID"], row["Summed 总需求"]), axis=1
    )
    reaction_df["  "] = None
    reaction_df.reset_index()
    material_ids = reaction_df["Unique Material ID"].tolist()
    min_sell_prices = ESI.get_jitasell(material_ids)
    # Assign prices
    reaction_df["Min Jita Sell"] = reaction_df["Unique Material ID"].map(min_sell_prices)

    reaction_df = reaction_df.reset_index(drop=True)
    for i, (idx, row) in enumerate(reaction_df.iterrows(), start=2):
        if pd.isna(row["Rounded 总需求"]) or str(row["Rounded 总需求"]).strip() == "":
            continue  # skip this row

        existing_formula = str(row["Rounded 总需求"])
        reaction_df.at[idx, "Total Cost"] = f"=({existing_formula[1:]})*(ReactionFormulas{count}!R{i})*({price_control})"

    reaction_df.at[0, "Total Cost Sum"] = f"=SUM(S{2}:S{len(reaction_df) + 1})"

    # Filter out rows where Total Cost is not empty or zero
    valid_rows = reaction_df[reaction_df["Total Cost"].notna() & (reaction_df["Total Cost"] != 0)]

    # Initialize 已完成数量 column
    reaction_df["已完成数量"] = ""
    start_row = 2
    # Fill values only for valid rows
    for i, idx in enumerate(valid_rows.index, start=start_row):
        reaction_df.at[idx, "剩余数量"] = f'=P{i}-U{i}'
        reaction_df.at[idx, "已完成数量"] = 0


    #get EVI
    reaction_df["   "] = None
    reaction_df["    "] = None
    reaction_df["     "] = None
    reaction_df.reset_index()
    material_ids = reaction_df["Unique Material ID"].tolist()
    min_sell_prices = ESI.get_adjusted_prices(material_ids)
    # Assign prices
    reaction_df["Adjusted Price"] = reaction_df["Unique Material ID"].map(min_sell_prices)

    reaction_df = reaction_df.reset_index(drop=True)
    for i, (idx, row) in enumerate(reaction_df.iterrows(), start=2):
        if pd.isna(row["Rounded 总需求"]) or str(row["Rounded 总需求"]).strip() == "":
            continue  # skip this row

        existing_formula = str(row["Rounded 总需求"])
        reaction_df.at[idx, "Total Adjusted Cost"] = f"=({existing_formula[1:]})*ReactionFormulas{count}!Z{i}"
    reaction_df.at[0, "Total Adjusted Cost Sum"] = f"=SUM(ReactionFormulas{count}!AA{2}:ReactionFormulas{count}!AA{len(reaction_df) + 1})"

    return reaction_df, 0

def generate_blueprint_excel(name,item_ID, runs, price_control):

    #Sheet1
    start_row = 2
    #info page
    info_df = pd.DataFrame({
        "物品": name,
        "计划生产": [runs] * len(name)
    })
    #Sheet2
    filtered_df = blueprints.loc[item_ID]
    #filter material ID
    # if "Material ID" in filtered_df.columns:
    #     filtered_df = filtered_df.drop(columns=["Material ID"])
    #number of production
    filtered_df["材料效率"] = 1
    filtered_df.rename(columns={"Material EN": "Unique Material EN"}, inplace=True)
    filtered_df["Rounded 总需求"] = [
        f"=ROUNDUP(ComponentBlueprint!F{start_row + i}*ComponentBlueprint!G{start_row + i}, 0)*Info!B2" for i in range(len(filtered_df))
    ]
    filtered_df[""] = None
    filtered_df["已完成数量"] = 0
    filtered_df["剩余数量"] = [f"=H{start_row+i}-J{start_row+i}" for i in range(len(filtered_df))]

    filtered_df[" "] = None
    material_ids = filtered_df["Material ID"].tolist()
    min_sell_prices = ESI.get_jitasell(material_ids)
    # Assign prices
    filtered_df["material min jita sell"] = filtered_df["Material ID"].map(min_sell_prices)

    #get EVI
    filtered_df["   "] = None
    filtered_df["    "] = None
    filtered_df["     "] = None
    filtered_df.reset_index()
    material_ids = filtered_df["Material ID"].tolist()
    min_sell_prices = ESI.get_adjusted_prices(material_ids)
    # Assign prices
    filtered_df["Adjusted Price"] = filtered_df["Material ID"].map(min_sell_prices)

    filtered_df = filtered_df.reset_index(drop=True)
    for i, (idx, row) in enumerate(filtered_df.iterrows(), start=2):
        if pd.isna(row["Rounded 总需求"]) or str(row["Rounded 总需求"]).strip() == "":
            continue  # skip this row

        existing_formula = str(row["Rounded 总需求"])
        filtered_df.at[idx, "Total Adjusted Cost"] = f"=({existing_formula[1:]})*ComponentBlueprint!Q{i}"
    filtered_df.at[0, "Total Adjusted Cost Sum"] = f"=SUM(ComponentBlueprint!R{2}:ComponentBlueprint!R{len(filtered_df) + 1})"

    #generate bluprint sheet
    blueprint_sheets = []
    count = 1
    component_df, check = generate_blueprint_sheet(filtered_df,count,price_control)
    while check != 1:
        count += 1
        blueprint_sheets.append(component_df)
        component_df, check = generate_blueprint_sheet(component_df,count,price_control)
    total_sheets = [info_df,filter]
    total_sheets.extend(blueprint_sheets)

    #start of reaction sheet
    material_totals = {}  # Dictionary to store aggregated results
    #get material from previous sheet
    for sheet in total_sheets:  # total_sheets is a list of DataFrames
        if isinstance(sheet, pd.DataFrame):
            if "Unique Material EN" in sheet.columns and "Rounded 总需求" in sheet.columns:
                for _, row in sheet.iterrows():
                    material = row["Unique Material EN"]
                    amount = row["Rounded 总需求"]
                    if material in material_totals:
                        material_totals[material] = fix_formulas(f"{material_totals[material]}{amount}")
                    else:
                        material_totals[material] = amount
    #creating dummy sheet sum up all materila needs from bluprint
    material_ids = list(material_totals.keys())
    reactions_prepare_df = items.loc[items["en_name"].isin(material_ids), ["en_name"]]
    reactions_prepare_df.rename(columns={"en_name": "Unique Material EN"}, inplace=True)
    reactions_prepare_df["Material ID"] = reactions_prepare_df.index
    # Create a new column based on material_totals dictionary
    reactions_prepare_df["Rounded 总需求"] = reactions_prepare_df["Unique Material EN"].map(material_totals)
    #loop for reaction sheet
    reaction_sheets = []
    count = 1
    reaction_df, check = generate_reaction_sheet(reactions_prepare_df,count, price_control)
    while check != 1:
        count += 1
        reaction_sheets.append(reaction_df)
        reaction_df, check = generate_reaction_sheet(reaction_df,count,price_control)


    info_df[" "] = None

    #sum up adjust price
    reaction_EIV = "="
    count = 1
    for sheet in reaction_sheets:
        cell_ref = f"ReactionFormulas{count}!AB2+"
        reaction_EIV += cell_ref
        count +=1
    reaction_EIV = reaction_EIV[:-1]
    info_df.loc[0, "Raw Reaction EIV"] = reaction_EIV
    #blueprint
    count = 1
    material_EIV = "="
    for sheet in blueprint_sheets:
        cell_ref = f"MaterialBlueprints{count}!AA2+"
        material_EIV += cell_ref
        count += 1
    material_EIV = material_EIV[:-1]
    info_df.loc[0, "Raw Material EIV"] = material_EIV
    info_df.loc[0, "Raw Component EIV"] = "=ComponentBlueprint!S2"

    #set new row for info
    required_rows = 10
    if len(info_df) < required_rows:
        # Create the missing rows with None values
        missing_rows = required_rows - len(info_df)
        new_rows = pd.DataFrame([ [None] * len(info_df.columns) ] * missing_rows, columns=info_df.columns)

        # Append the new rows
        info_df = pd.concat([info_df, new_rows], ignore_index=True)

    info_df["   "] = None
    info_df["  "] = None
    info_df.iloc[0:3, info_df.columns.get_loc('  ')] = ["星系成本", "scc", "tax"]
    info_df["反应税率"] =  None
    info_df.iloc[0:3, info_df.columns.get_loc('反应税率')] = [0.02, 0.04, 0.025]
    info_df["组件税率"] = None
    info_df.iloc[0:3, info_df.columns.get_loc('组件税率')] = [0.02, 0.04, 0.025]
    info_df["组装税率"] = None
    info_df.iloc[0:3, info_df.columns.get_loc('组装税率')] = [0.02, 0.04, 0.025]

    info_df.loc[0, "Reaction EIV"] = "=SUM(I2:I4)*D2"
    info_df.loc[0, "Material EIV"] = "=SUM(J2:J4)*E2"
    info_df.loc[0, "Component EIV"] = "=SUM(K2:K4)*F2"
    info_df.loc[0, "Total Tax"] = "=L2+M2+N2"


    # Write the filtered rows into a new Excel file
    output_file = f"blueprint_{item_ID}.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        info_df.to_excel(writer, index=False, sheet_name="Info")
        filtered_df.to_excel(writer, index=False, sheet_name="ComponentBlueprint")
        for i in range(len(blueprint_sheets)):
            blueprint_sheets[i].to_excel(writer, index=False, sheet_name=f"MaterialBlueprints{i+1}")
        for i in range(len(reaction_sheets)):
            reaction_sheets[i].to_excel(writer, index=False, sheet_name=f"ReactionFormulas{i+1}")
        # reactions_prepare_df.to_excel(writer, index=False, sheet_name="ReactionPrepareBlueprint")
        # reaction_df.to_excel(writer, index=False, sheet_name="ReactionFormulas1")


    print(f"✅ Excel generated: {output_file}")

def name_to_id(names):

    matched_items = items[items["zh_name"].isin(names)]
    # Step 3: Get corresponding Item Path IDs
    ids = matched_items["id"].unique().tolist()

    return ids



# # Example usage
if __name__ == "__main__":


    names = ["泽尼塔级蓝图"]
    ids = name_to_id(names)
    items.set_index("id", inplace=True)
    runs = 10
    price_control = 1.0
    # item_id = ESI.bulk_names_to_ids(names)
    generate_blueprint_excel(names,ids,runs, price_control)


    end_time = time.time()  # End timer
    print(f"Execution time: {end_time - start_time:.2f} seconds")
