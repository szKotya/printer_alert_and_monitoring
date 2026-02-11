test = 'РЕЖИМ ОЖИДАНИЯ...'
print(test)

test = test.lower()
test = test[0].upper() + test[1:]
test = test.replace(".", "")
print(test)