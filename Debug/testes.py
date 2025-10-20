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
        # column_match = r'column\s+([A-Za-z0-9_]+)'
        #columns = re.findall(column_match, tmdl_data)
        columns = re.findall(r'(?mi)^\s*column\s+([A-Za-z0-9_]+)\s*$', tmdl_data)

        # Extract mode (Import/DirectQuery)
        # mode_match = re.search(r'(?mi)^\s*mode:\s*([^\r\n]+)', tmdl_data)
        # mode = mode_match.group(1).strip() if mode_match else None
        mode_match = re.search(r'(?mi)^\s*mode\s*:\s*([^\r\n]+)', tmdl_data)
        mode = mode_match.group(1).strip() if mode_match else (
            re.search(r'(?mi)^\s*annotation\s+PBI_DataMode\s*=\s*"?(.*?)"?\s*$', tmdl_data).group(1).strip()
            if re.search(r'(?mi)^\s*annotation\s+PBI_DataMode', tmdl_data) else None
        )
        if mode and ((mode.startswith("```") and mode.endswith("```")) or (mode.startswith("`") and mode.endswith("`"))):
            mode = mode.strip("`").strip()

        # Extract queryGroup (if any)
        # query_group_match = re.search(r'(?mi)^\s*queryGroup:\s*([^\r\n]+)', tmdl_data)
        # query_group = query_group_match.group(1).strip() if query_group_match else None
        query_group_match = re.search(r'(?mi)^\s*queryGroup\s*:\s*([^\r\n]+)', tmdl_data) \
            or re.search(r'(?mi)^\s*queryGroup\s+([^\r\n]+)', tmdl_data)
        query_group = query_group_match.group(1).strip() if query_group_match else None

        # Extract M code
        # expression_match = re.search(
        #     r'(?ms)^\s*source\s*=\s*\n'       # find the start of "source ="
        #     r'((?:[ \t]+.*\n)+?)'             # capture indented lines
        #     r'(?=^[^\t ]|^\s*annotation|\Z)', # stop when indentation ends or next annotation
        #     tmdl_data
        # )
        # m_query = expression_match.group(1) if expression_match else None
        
        def unescape_quoted(s: str) -> str:
            try:
                return bytes(s, "utf-8").decode("unicode_escape")
            except Exception:
                return s

        def extract_m_code(tmdl_text: str):
            # A) Quoted expression form: expression = "let\n...."
            mq = re.search(r'(?ms)^\s*expression\s*=\s*"((?:[^"\\]|\\.)*)"', tmdl_text)
            if mq:
                return unescape_quoted(mq.group(1)).strip()

            # B) Indented block after a standalone "source =" line
            src = re.search(r'(?m)^\s*source\s*=\s*$', tmdl_text)
            if not src:
                # Fallback: inline source = <content> ... up to next annotation or EOF
                m_inline = re.search(r'(?ms)^\s*source\s*=\s*(.+?)(?=^\s*annotation\b|^\S|\Z)', tmdl_text)
                return m_inline.group(1).rstrip() if m_inline else None

            start = src.end()
            lines = tmdl_text[start:].splitlines()

            # find first non-empty line to set base indentation
            i = 0
            while i < len(lines) and lines[i].strip() == "":
                i += 1
            if i == len(lines): 
                return None

            first = lines[i]
            base_indent = len(first) - len(first.lstrip())

            out = []
            for line in lines[i:]:
                stripped = line.lstrip()
                # stop if an annotation starts (even if indented)
                if stripped.startswith("annotation "):
                    break
                # stop if indentation drops below the first M line
                if stripped and (len(line) - len(stripped)) < base_indent:
                    break
                out.append(line)

            # trim trailing blank lines
            while out and out[-1].strip() == "":
                out.pop()

            return "\n".join(out)

        m_code = extract_m_code(tmdl_data)

        tables_data[table_name] = {
            "columns": columns,
            "import_mode": mode,
            "query_group": query_group,
            "m_code": m_code
        }

# Show the final json
print(json.dumps(tables_data, indent=4))