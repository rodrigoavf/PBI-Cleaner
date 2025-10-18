import os, json

root_dir = "C:/Users/rodrigo.ferreira/Desktop/Devoteam/Supply & Purchasing.pbip"

# From root_dir remove the file extension, add .SemanticModel and a subfolder called DAXQueries
folder_of_interest = os.path.splitext(root_dir)[0] + ".SemanticModel/DAXQueries"
json_path = folder_of_interest + "/.pbi/daxQueries.json"

# Create a variable that holds the dax queries names, which are the values in the json_dax_queries file in the "tabOrder" key
with open(json_path, "r", encoding="utf-8") as f:
    json_content = json.load(f)
dax_query_names = json_content["tabOrder"]
default_query_name = json_content["defaultTab"]

queries = {}
for query in dax_query_names:
    with open(folder_of_interest + "/" + query + ".dax", "r", encoding="utf-8") as f:
        dax_code = f.read()
    queries[query] = dax_code

print(queries)

# Write back to the json file with updated tab order and default tab
# new_tab_order = ["Query 2", "Query 1", "Query 3"]
# new_default_tab = "Query 2"
# updated_data = {
#     "tabOrder": new_tab_order,
#     "defaultTab": new_default_tab
# }
# with open(json_path, "w", encoding="utf-8") as f:
#     json.dump(updated_data, f, indent=4)

# # Show DAX query code for selected query name
# selected_query_name = "Query 1"

# # Open and read the contents of a .dax file
# with open(folder_of_interest + "/" + selected_query_name + ".dax", "r", encoding="utf-8") as f:
#     dax_code = f.read()

# print(dax_code)