import os, re, ast, json

root_dir = "C:/Users/rodrigo.ferreira/Desktop/Devoteam/Supply & Purchasing.pbip"

model_tmdl = os.path.splitext(root_dir)[0] + ".SemanticModel/definition/model.tmdl"

with open(model_tmdl, 'r', encoding='utf-8') as file:
    model_tmdl_data = file.read()

# Extract PBI_QueryOrder
pattern = r"annotation\s+PBI_QueryOrder\s*=\s*(\[.*?\])"
match = re.search(pattern, model_tmdl_data, re.DOTALL)
if match:
    list_str = match.group(1)
    query_order = ast.literal_eval(list_str)
    print(query_order)

# Extract all query groups and their orders
pattern = r"queryGroup\s+(\w+)\s*\n\s*annotation\s+PBI_QueryGroupOrder\s*=\s*(\d+)"
matches = re.findall(pattern, model_tmdl_data)
query_groups = {name: int(order) for name, order in matches}
print(query_groups)

# For each .tmdl file in tables_tmdl open the file and fill a json
# The main key will be the table name extracted from the file name
# Followed by a list of columns extracted from each file
# Then the import mode extracted from the model.tmdl file
# The query group if applicable
# The M code if applicable
tables_tmdl = os.path.splitext(root_dir)[0] + ".SemanticModel/definition/tables"
tables_data = {}
for filename in os.listdir(tables_tmdl):
    if filename.endswith(".tmdl"):
        table_name = os.path.splitext(filename)[0]
        table_path = os.path.join(tables_tmdl, filename)
        with open(table_path, 'r', encoding='utf-8') as file:
            tmdl_data = file.read()
        
        # Extract all column names
        column_match = r'column\s+([A-Za-z0-9_]+)'
        columns = re.findall(column_match, tmdl_data)
        
        # Extract mode (Import/DirectQuery)
        mode_match = re.search(r'(?mi)^\s*mode:\s*([^\r\n]+)', tmdl_data)
        mode = mode_match.group(1).strip() if mode_match else None
        
        # Extract queryGroup (if any)
        query_group_match = re.search(r'(?mi)^\s*queryGroup:\s*([^\r\n]+)', tmdl_data)
        query_group = query_group_match.group(1).strip() if query_group_match else None
        
        # Extract M code
        expression_match = re.search(
            r'(?ms)^\s*source\s*=\s*\n'           # the "source =" line
            r'((?:[ \t].*\n)+?)'                  # the indented M code lines
            r'\n(?=\s*annotation|\Z)',            # stop before next annotation or end-of-file
            tmdl_data
        )
        m_query = expression_match.group(1) if expression_match else None
        
        tables_data[table_name] = {
            "columns": columns,
            "import_mode": mode,
            "query_group": query_group,
            "m_code": m_query
        }

# Show the final json
print(json.dumps(tables_data, indent=4))