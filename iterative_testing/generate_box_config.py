# arguments
# python3 <programname> <number of boxes> <start column>

def generateNextAlphabet(current):
  if (len(current) == 1):
    if (current == "Z"):
      return "AA"
    else:
      return chr(ord(current) + 1)
  else:
    lastIndex = len(current)
    currentIndex = lastIndex - 1
    returnString = current
    letterToModify = (current)[currentIndex]

    while (letterToModify == "Z" and currentIndex >= 0):
      letterToModify = (current)[currentIndex]
      returnString = returnString[0:currentIndex] + "A" + returnString[currentIndex + 1:lastIndex]
      currentIndex = currentIndex - 1
    return returnString[:currentIndex] + chr(ord(letterToModify) + 1) + returnString[currentIndex+1:lastIndex]


print(generateNextAlphabet("AZZ"))
print(generateNextAlphabet("Z"))
print(generateNextAlphabet("BAA"))


