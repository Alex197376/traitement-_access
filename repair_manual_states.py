import json
import os

path_in = "manual_states.json"
path_out = "manual_states_repair.json"

with open(path_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Tentative simple de réparation
content = content.replace("'", '"')  # guillemets simples vers doubles

lines = content.splitlines()
cleaned_lines = []

for line in lines:
    try:
        # test si la ligne est un JSON valide (encapsulée comme un objet)
        json.loads("{" + line.strip().strip(',') + "}")
        cleaned_lines.append(line.strip().strip(','))
    except:
        continue

# Réécriture dans un nouveau fichier
final_content = "{\n" + ",\n".join(cleaned_lines) + "\n}"

try:
    data = json.loads(final_content)
    with open(path_out, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    print(f"✅ Réparation réussie : {path_out}")
except Exception as e:
    print(f"❌ Impossible de réparer : {e}")
