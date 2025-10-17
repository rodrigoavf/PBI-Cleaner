##############################################################################################################################
# PARA QUE SERVE ESTE SCRIPT
# Este script é utilizado para renomear measures em ficheiros .json e .tmdl dentro de uma pasta e suas subpastas.
# Ele acelera o processo de renomeação em lote, permitindo que você substitua chaves específicas por novos nomes de forma rápida e eficiente.
# Além de previnir erros de apontamento que podem acontecer quando se renomeia measures utilizando o Tabular Editor.

# INSTRUÇÕES DE USO
# 1 - Salve o report de PBI como .pbip com as opções de TMDL e PBIR ativadas.
# 2 - Salve este script na mesma pasta do ficheiro .pbip
# 3 - Atualize o dicionário `find_replace_dict` com as chaves e valores que deseja substituir.
# 4 - Feche o report PBI se estiver aberto.
# 5 - Execute o script. Ele irá procurar por todos os ficheiros .json e .tmdl na pasta e subpastas, e substituir as chaves pelos valores correspondentes.
# 6 - Aguarde pela mensagem de conclusão do script.
# 7 - Volte a abrir o report PBI usando o ficheiro .pbip e verifique se as alterações foram aplicadas corretamente.
# 8 - Se quiser, volte a salvar o report PBI como .pbix ou qualquer outro formato desejado.
##############################################################################################################################

import os
import re

# Root directory to start searching
root_dir = os.path.dirname(os.path.abspath(__file__))

# Dictionary for batch find and replace: {find_str: replace_str}
find_replace_dict = {
    "CY_CUSTO_MEDIO_STOCKS":"KPI-003.03 - CY",
    "PY_CUSTO_MEDIO_STOCKS":"KPI-003.03 - PY",
    "V_EXEC_CustoMedioStocks":"KPI-003.03 - V_EXEC",
    "V_HMLG_CustoMedioStocks":"KPI-003.03 - V_HMLG",
    "ΔPY_CUSTO_MEDIO_STOCKS":"KPI-003.03 - ΔPY",

    "CY_STOCKS_QTY":"KPI-003.01 - CY",
    "PY_STOCKS_QTY":"KPI-003.01 - PY",
    "V_EXEC_StocksQtd":"KPI-003.01 - V_EXEC",
    "V_HMLG_StocksQtd":"KPI-003.01 - V_HMLG",
    "ΔPY_STOCKS_QTY":"KPI-003.01 - ΔPY",
    "Δ%PY_STOCKS_QTY":"KPI-003.01 - Δ%PY",

    "CY_STOCKS_SUPPLIER_VALOR":"KPI-003.02 D - CY",
    "PY_STOCKS_SUPPLIER_VALOR":"KPI-003.02 D - PY",
    "Δ%PY_STOCKS_SUPPLIER_VALOR":"KPI-003.02 D - Δ%PY",
    "ΔPY_STOCKS_SUPPLIER_VALOR":"KPI-003.02 D - ΔPY",

    "CY_STOCKS_VALOR":"KPI-003.02 - CY",
    "PY_STOCKS_VALOR":"KPI-003.02 - PY",
    "V_EXEC_StocksValor":"KPI-003.02 - V_EXEC",
    "V_HMLG_StocksValor":"KPI-003.02 - V_HMLG",
    "Δ%PY_STOCKS_VALOR":"KPI-003.02 - Δ%PY",
    "ΔPY_STOCKS_VALOR":"KPI-003.02 - ΔPY",

    "CY %Stock Available Per Site - Inventory Value":"KPI-034.02 - CY",
    "PY %Stock Available Per Site - Inventory Value":"KPI-034.02 - PY",
    "%PY %Stock Available Per Site - Inventory Value":"KPI-034.02 - %PY",
    "Δ%PY %Stock Available Per Site - Inventory Value":"KPI-034.02 - Δ%PY",
    "ΔPY %Stock Available Per Site - Inventory Value":"KPI-034.02 - ΔPY",

    "CY Stock Available Per Site - Inventory Value":"KPI-034.01 - CY",
    "PY Stock Available Per Site - Inventory Value":"KPI-034.01 - PY",
    "%PY Stock Available Per Site - Inventory Value":"KPI-034.01 - %PY",
    "Δ%PY Stock Available Per Site - Inventory Value":"KPI-034.01 - Δ%PY",
    "ΔPY Stock Available Per Site - Inventory Value":"KPI-034.01 - ΔPY",

    "CY Stock Available Per Site - Inventory Weight":"KPI-034.03 - CY",
    "PY Stock Available Per Site - Inventory Weight":"KPI-034.03 - PY",
    "%PY Stock Available Per Site - Inventory Weight":"KPI-034.03 - %PY",
    "Δ%PY Stock Available Per Site - Inventory Weight":"KPI-034.03 - Δ%PY",
    "ΔPY Stock Available Per Site - Inventory Weight":"KPI-034.03 - ΔPY",

    "CY %Stock Available Per Site - Inventory Weight":"KPI-034.04 - CY",
    "PY %Stock Available Per Site - Inventory Weight":"KPI-034.04 - PY",
    "%PY %Stock Available Per Site - Inventory Weight":"KPI-034.04 - %PY",
    "Δ%PY %Stock Available Per Site - Inventory Weight":"KPI-034.04 - Δ%PY",
    "ΔPY %Stock Available Per Site - Inventory Weight":"KPI-034.04 - ΔPY",

    "CY_VENDAS_LÍQUIDAS_VALOR":"KPI-002.01 - CY",
    "PY_VENDAS_LÍQUIDAS_VALOR":"KPI-002.01 - PY",
    "Δ%PY_VENDAS_LÍQUIDAS_VALOR":"KPI-002.01 - Δ%PY",
    "ΔPY_VENDAS_LÍQUIDAS_VALOR":"KPI-002.01 - ΔPY",

    "KPI-035.03 - CY Análise Performance Stock":"KPI-035.03 - CY",
    "PY Análise Performance Stock":"KPI-035.03 - PY",
    "%PY Análise Performance Stock":"KPI-035.03 - %PY",
    "Δ%PY Análise Performance Stock":"KPI-035.03 - Δ%PY",
    "ΔPY Análise Performance Stock":"KPI-035.03 - ΔPY"
}


def token_pattern(key):
    """
    Match KEY only when it’s a whole token delimited by start/end, spaces, [], or single-quote.
    """
    prefix = r"(?:(?<=^)|(?<=[\s\[\]']))"
    suffix = r"(?:(?=$)|(?=[\s\[\]']))"
    return re.compile(prefix + re.escape(key) + suffix)

# Pre-compile all the regexes for token-only replacement
replacement_patterns = {
    key: (token_pattern(key), val)
    for key, val in find_replace_dict.items()
}

def apply_visual_context(text):
    """
    In visual.json only: replace .KEY" → .VALUE
                       and "KEY"  → "VALUE"
    regardless of spaces in VALUE.
    """
    for key, val in find_replace_dict.items():
        # .KEY"
        pat_dot = re.compile(r'\.' + re.escape(key) + r'(?=")')
        text = pat_dot.sub('.' + val, text)
        # "KEY"
        pat_quote = re.compile(r'"' + re.escape(key) + r'"')
        text = pat_quote.sub(f'"{val}"', text)
    return text

def apply_measure_quotes(text):
    """
    For any key without spaces, if it's preceded by 'measure ' (as its own token)
    and its replacement contains spaces, wrap it in single quotes.
    """
    prefix = r"(?:(?<=^)|(?<=[\s\[\]']))"
    suffix = r"(?:(?=$)|(?=[\s\[\]']))"
    for key, val in find_replace_dict.items():
        if ' ' not in key:
            pat = re.compile(prefix + r"measure\s+" + re.escape(key) + suffix)
            def _repl(m, val=val):
                if ' ' in val:
                    return f"measure '{val}'"
                else:
                    return f"measure {val}"
            text = pat.sub(_repl, text)
    return text

def process_file(path, is_json: bool):
    with open(path, 'r', encoding='utf-8') as f:
        orig = f.read()

    new = orig

    # 1) If this is visual.json, apply the dot-KEY" / "KEY" rules first
    if os.path.basename(path) == 'visual.json':
        new = apply_visual_context(new)

    # 2) Apply the special 'measure …' quoting rules
    new = apply_measure_quotes(new)

    # 3) Apply only whole-token replacements
    for key, (pat, val) in replacement_patterns.items():
        new = pat.sub(val, new)

    if new != orig:
        with open(path, 'w', encoding='utf-8') as f:
            f.write(new)
        kind = "JSON" if is_json else "TMDL"
        print(f"  ✅ Replaced in {kind}: {path}")
    else:
        kind = "JSON" if is_json else "TMDL"
        print(f"  ⚠️ No matches in {kind}: {path}")

# File extensions to include
valid_extensions = {".json", ".tmdl"}

print("CWD:", os.getcwd())
print("Walking from:", os.path.abspath(root_dir))

for dirpath, dirnames, filenames in os.walk(root_dir, topdown=True, followlinks=False):
    print("Scanning:", dirpath)
    for fn in filenames:
        ext = os.path.splitext(fn)[1].lower()
        if ext not in valid_extensions:
            continue
        full = os.path.join(dirpath, fn)
        print(" • Found file:", full)
        process_file(full, is_json=(ext == ".json"))

print("\n✨ All done! Scanning and replacements complete.")