# -*- coding: utf-8 -*-
"""
Select2 (다중선택 드롭다운) 입력 통합
"""

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time


class Select2Handler:
    """Select2 필드 입력 통합"""
    
    def __init__(self, driver, wait_time=0.25):
        self.driver = driver
        self.wait_time = wait_time
    
    def fill(self, 
            ul_selector: str,
            input_selector: str,
            values: list,
            clear_existing: bool = True) -> bool:
        """Select2 필드에 값들 입력"""
        try:
            if clear_existing:
                self.clear(ul_selector)
            
            for v in values:
                if v:
                    self.add(input_selector, v)
            
            return True
        except Exception as e:
            print(f"❌ Select2 입력 실패: {e}")
            return False
    
    def clear(self, ul_selector: str) -> bool:
        """Select2 필드의 모든 선택값 제거"""
        try:
            ul = self.driver.find_element(By.CSS_SELECTOR, ul_selector)
        except:
            return False
        
        while True:
            buttons = ul.find_elements(By.CSS_SELECTOR, "li span, li i.fa")
            
            if not buttons:
                break
            
            for btn in buttons:
                try:
                    self.driver.execute_script("arguments[0].click();", btn)
                    time.sleep(0.1)
                except:
                    pass
        
        return True
    
    def add(self, input_selector: str, value: str) -> bool:
        """Select2 필드에 값 하나 추가"""
        try:
            inp = self.driver.find_element(By.CSS_SELECTOR, input_selector)
            inp.clear()
            inp.send_keys(value)
            time.sleep(self.wait_time)
            inp.send_keys(Keys.ENTER)
            time.sleep(self.wait_time)
            return True
        except Exception as e:
            print(f"⚠ 값 추가 실패({value}): {e}")
            return False