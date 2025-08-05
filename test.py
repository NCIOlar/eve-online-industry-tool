import yaml
import pandas as pd
import getID
import ESI
import time



start_time = time.time()

#Sheet3
# material_names = items.loc[material_ids, "en_name"]  # Map Material ID -> Material Name
# blueprint_names = material_names + " Blueprint"
# component_df = blueprints[blueprints["Blueprint EN"].isin(blueprint_names)].copy()
# component_df["材料效率"] = 1
#
# component_df["Base Name"] = component_df["Blueprint EN"].str.replace(" Blueprint", "", regex=False)
# # Create a mapping from Material EN to 总需求
# mapping = dict(zip(filtered_df["Material EN"], filtered_df["总需求"]))
# component_df["总需求"] = component_df["Base Name"].map(mapping)
#
# component_df = component_df.reset_index(drop=True)
# for i, (idx, row) in enumerate(component_df.iterrows(), start=2):  # start=2 for Excel rows
#     existing_formula = str(row["总需求"])  # convert to string to avoid issues
#     component_df.at[idx, "总需求"] = f"={existing_formula[1:]}*ROUNDUP(F{i}*G{i},0)"
# #Drop base name helper colume
# component_df = component_df.drop(columns=["Base Name"])



# # Example usage
names = ["伊什塔级蓝图"]
item_id = ESI.bulk_names_to_ids(names)
generate_blueprint_excel(item_id)

end_time = time.time()  # End timer
print(f"Execution time: {end_time - start_time:.2f} seconds")
