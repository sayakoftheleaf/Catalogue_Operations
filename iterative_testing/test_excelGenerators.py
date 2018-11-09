import generate_box_config
import pytest

 
def test_single_letter_not_end_of_alphabet ():
  assert generate_box_config.generateExcelColumns("M") == "N"
  
def test_single_letter_end_of_alphabet ():
  assert generate_box_config.generateExcelColumns("Z") == "AA"

def test_multiple_letters_all_not_end_of_alphabet ():
  assert generate_box_config.generateExcelColumns("FG") == "FH"

def test_multiple_letters_lastchar_end_of_alphabet ():
  assert generate_box_config.generateExcelColumns("AZ") == "BA"

def test_multiple_letters_nonlastchar_end_of_alphabet ():
  assert generate_box_config.generateExcelColumns("BZX") == "BZY"

def test_multiple_letters_all_end_of_alphabet ():
  assert generate_box_config.generateExcelColumns("ZZZ") == "AAAA"
  
