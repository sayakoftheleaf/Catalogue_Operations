import box_order
import pytest
import box_order

class TestBoxOrder(object):
  firstResultDict = {'L': 0, 'B': 1, 'H': 2,'W': 3}
  secondResultDict = {'L': 3, 'B': 0, 'H': 2,'W': 1}
 
  def test_different_orders_UpperCase(self):
    assert cmp(box_order.mapBoxes("L,B,H,W"), self.firstResultDict) == 0
    assert cmp(box_order.mapBoxes("B,W,H,L"), self.secondResultDict) == 0
  
  def test_different_orders_LowerCase(self):
    assert cmp(box_order.mapBoxes("l,b,h,w"), self.firstResultDict) == 0
    assert cmp(box_order.mapBoxes("b,w,h,l"), self.secondResultDict) == 0
 
  def test_different_orders_UpperCaseWithSpace(self):
    assert cmp(box_order.mapBoxes("L, B, H, W"), self.firstResultDict) == 0
    assert cmp(box_order.mapBoxes("B, W, H, L"), self.secondResultDict) == 0
  
  def test_different_orders_LowerCaseWithSpace(self):
    assert cmp(box_order.mapBoxes("l, b, h, w"), self.firstResultDict) == 0
    assert cmp(box_order.mapBoxes("b, w, h, l"), self.secondResultDict) == 0
             
