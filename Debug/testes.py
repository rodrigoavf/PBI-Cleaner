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
            table_data = file.read()
        
        # Extract columns
        # The pattern is Column ColumnName (Might be with quotes, and may be followed a space followed by a = sign)
        # Example: column DESCRIPTION_PROG
        # Example 2: column DESCRIPTION_PROG =
        column_pattern = r"column\s+['\"]?(\w+)['\"]?\s*(?:=|$)"
        columns = re.findall(column_pattern, table_data)
        
        # Extract import mode
        # The pattern in mode: importMode
        # Exmaple: mode: import
        import_mode_pattern = r"mode:\s+(\w+)"
        import_mode_match = re.search(import_mode_pattern, model_tmdl_data)
        import_mode = import_mode_match.group(1) if import_mode_match else "Unknown"
        
        # Extract query group
        # The pattern is queryGroup: queryGroupName
        # Example: queryGroup: Sales
        query_group_pattern = r"queryGroup:\s+(\w+)"
        query_group_match = re.search(query_group_pattern, table_data)
        query_group = query_group_match.group(1) if query_group_match else "Default"
        
        # Extract M code
        # The pattern is source = (the code can be multiline until a line before an empty line or the end of the file)
        # Exemple:  
            # source =
                # let
                #     Origem = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i44FAA==", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Coluna1 = _t]),
                #     #"Tipo Alterado" = Table.TransformColumnTypes(Origem,{{"Coluna1", type text}}),
                #     #"Removed Columns" = Table.RemoveColumns(#"Tipo Alterado",{"Coluna1"})
                # in
                #     #"Removed Columns"
        m_code_pattern = r"source\s*=\s*(let\s.*?)(?:\n\s*\n|$)"
        m_code_match = re.search(m_code_pattern, table_data, re.DOTALL)
        m_code = m_code_match.group(1).strip() if m_code_match else ""
        
        tables_data[table_name] = {
            "columns": columns,
            "import_mode": import_mode,
            "query_group": query_group,
            "m_code": m_code
        }

# Show the final json
print(json.dumps(tables_data, indent=4))