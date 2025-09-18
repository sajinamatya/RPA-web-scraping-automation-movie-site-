import os
import json

def read_vaultfile():
    with open(os.path.join('.', 'vault', 'data.json'), 'r') as file:
        data = json.load(file)
        print(data)
        return data