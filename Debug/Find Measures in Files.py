##############################################################################################################################
# PARA QUE SERVE ESTE SCRIPT
# Este script é utilizado para procurar por uma chave específica em todos os ficheiros .json e .tmdl dentro de uma pasta e suas subpastas.
# Ele é útil para identificar onde uma determinada measure é utilizada, facilitando a manutenção e atualização.

# INSTRUÇÕES DE USO
# 1 - Salve o report de PBI como .pbip com as opções de TMDL e PBIR ativadas.
# 2 - Salve este script na mesma pasta do ficheiro .pbip
# 3 - Atualize a variável `target` com a chave que deseja encontrar.
# 4 - Execute o script. Ele irá procurar por todos os ficheiros .json e .tmdl na pasta e subpastas, e apontar quais possuem a chave procurada.
##############################################################################################################################

import os

# Root directory to start searching (you can change this or accept as an argument)
root_dir = os.path.dirname(os.path.abspath(__file__))

# The target substring to look for
target = "KPI-103.01 - CY"

# File extensions to include
valid_extensions = {".json", ".tmdl"}

def file_contains_target(path, target):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return target in f.read()
    except Exception as e:
        # Could log errors here if you like
        return False

def find_files_with_target(root_dir, target):
    matches = []
    for subdir, _, files in os.walk(root_dir):
        for fname in files:
            ext = os.path.splitext(fname)[1].lower()
            if ext in valid_extensions:
                full_path = os.path.join(subdir, fname)
                if file_contains_target(full_path, target):
                    matches.append(full_path)
    return matches

if __name__ == "__main__":
    hits = find_files_with_target(root_dir, target)
    if hits:
        print("Files containing", repr(target), ":")
        for p in hits:
            print(" -", p)
    else:
        print("No files found containing", repr(target))
