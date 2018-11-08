import toml

configDict = toml.load('config.toml')

sheet = configDict.get('sheet')

for key, value in sheet.items():
  print("key is " + str(key), "value is " + str(value))