from datetime import datetime
import os
import pathlib
import sys
import yaml

# print(os.path.dirname(__file__))
# print(os.path.abspath(os.path.dirname(__file__)))
# print(os.path.basename(__file__))

current_path = pathlib.Path(__file__).parent.resolve()
print(f'current_path : \n{current_path}')

# config_path = current_path._str+"\\..\\config"
# print(f'config_path : \n{config_path}')

# config_path = os.path.abspath(current_path._str+"\\Hoa\\Wolverine\\config")
# print(f'config_path : \n{config_path}')

# with open(os.path.abspath(f'{config_path}\\{"local_path.yaml"}'), 'r') as f:
#     config = yaml.load(f, Loader=yaml.SafeLoader)
# print(f'config : \n{config}')

# path = config.get("path")
# print(f'path : \n{path}')

current_file = os.path.basename(__file__)[:-3]
# print(f'current_file : \n{current_file}')

from func.read_config import read_config

path_config = "local_path.yaml"
path_config = read_config(path_config)
path = path_config.get("path")
# print(f'path : \n{path}')
current_file = path_config[os.path.basename(__file__)[:-3]]
print(f'current_file : \n{current_file}')
path = f"{path}\\{current_file}"
path = f"{current_path}\\..\\..\\..\\..\\..\\{path}"
path = os.path.abspath(path)
today_str = datetime.now().strftime("%m.%d.%Y")
path = f"{path}\\{today_str}"

print(f'path : \n{path}')
