import yaml

with open('data/config.yml', 'r', encoding='utf-8') as file:
    config = yaml.safe_load(file)


print(config)
