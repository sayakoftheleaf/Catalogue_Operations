# arguments
# python3 <programname> <number of boxes> <start column>

def generateExcelColumns (currentCol):
  lastIndex = len(currentCol) - 1
  currentIndex = lastIndex
  newCol = ''

  while (currentIndex >= 0):
    # find the next letter
    newLetter = findReplacement(currentCol[currentIndex]) 
    
    # if the next letter is an A, we need to do further computations
    if (newLetter == 'A'):
      # save the number of A's processed to be appended later
      newCol = newLetter + newCol
      
      # if the encountered A is not the first letter, then proceed
      if (currentIndex > 0):
        # move on to the next character on the left
        currentIndex = currentIndex - 1
        continue
      else: # if all the characters are A's, that means, we need to increase the length of the column name
        return 'A' + newCol
    else:
      # if we find a letter that is not A, we change that letter, 
      # and leave the letters left of it as is
      # and append all the previously processed characters, if any
      return modifyStringTill(currentCol, currentIndex, newLetter) + newCol

def modifyStringTill (originalString, indexValue, replacementValue):
  return originalString[: indexValue] + replacementValue

def findReplacement (letter):
  if (letter == 'Z'):
    return 'A'
  else:
    return chr(ord(letter) + 1)



