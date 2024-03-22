from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from time import sleep

class Chrome_Util:
    # コンストラクタ
    def __init__(self, driver_path, options_str_list=[], binary_location=None):
        # ChromeDriverの起動
        options = webdriver.ChromeOptions()
        if binary_location:
            options.binary_location = binary_location
        for option_str in options_str_list:
            options.add_argument(option_str)
        self._driver = webdriver.Chrome(service=Service(executable_path=driver_path), options=options)
        self.load_wait()
    
    # デストラクタ
    def __del__(self):
        self.close_driver

    # ドライバ取得
    @property
    def driver(self):
        return self._driver
    
    # URL取得
    @property
    def current_url(self):
        return self._driver.current_url

    # ドライバクローズ
    def close_driver(self):
        self._driver.close()
        self._driver.quit()
    
    # URLを開く
    def open_url(self, url, timeout=20):
        try:
            # ポップアップ無効
            self._driver.execute_script('window.onbeforeunload = function() {};')
            self._driver.get(url)
            self.load_wait(timeout)
            return True
        except:
            return False
    
    # ページ遷移
    def location_href(self, url):
        # ポップアップ無効
        self._driver.execute_script('window.onbeforeunload = function() {};')
        self._driver.execute_script(f'window.location.href = "{url}";')
        self.load_wait()
    
    # アラートを閉じる
    def close_alert(self, ele=None, timeout=5):
        # アラートを待機して取得し、閉じる
        try:
            if ele: ele.click()
            alert = WebDriverWait(self._driver, timeout).until(EC.alert_is_present())
            alert.dismiss()
            return True
        except:
            return False
    
    # 対象のエレメントまでスクロール
    def scrol(self, ele):
        try:
            self._driver.execute_script('arguments[0].scrollIntoView({ behavior: "smooth", block: "center" });', ele)
        except:
            pass

    # コンテンツが読み込まれるまで待機
    def load_wait(self, timeout=20, sleep_time=0):
        try:
            sleep(sleep_time)
            WebDriverWait(self._driver, timeout).until(EC.presence_of_all_elements_located)
        except Exception as e:
            print('タイムアウトしました')
            print(e)
    
    # 要素が使用可能になるまで待つlogin
    def implicitly_wait(self, timeout=20):
        self._driver.implicitly_wait(timeout)
    
    # 画面サイズ変更
    def set_window_size(self, width, height):
        self._driver.set_window_size(f'{width}', f'{height}')

    # javascript実行
    def exe_js(self, method_str, *args, timeout=20):
        result = self._driver.execute_script(f'return {method_str}', *args)
        self.load_wait(timeout)
        return result
    
    # リードオンリーのエレメントのリードオンリーを消す
    def remove_read_only(self, ele):
        self.exe_js("arguments[0].removeAttribute('readonly');", ele)
        return ele
    
    # エレメント削除
    def del_element(self, ele):
        self._driver.execute_script('arguments[0].remove();', ele)
    
    # エレメントのValue変更
    def set_ele_value(self, ele, val):
        self._driver.execute_script('arguments[0].value = arguments[1]', ele, val)
    
    # エレメントのValue取得
    def get_ele_value(self, ele):
        return self._driver.execute_script('return arguments[0].value', ele)
    
    # エレメント検索
    def find_element(self, mode, word, timeout=20):
        elements = self.find_elements(mode, word, timeout) or []
        return elements[0] if len(elements) > 0 else None

    # エレメント検索(複数)
    def find_elements(self, mode, word, timeout=20):
        tgt_mode = {
            'id'          : By.ID,
            'class'       : By.CLASS_NAME,
            'tag'         : By.TAG_NAME,
            'name'        : By.NAME,
            'title'       : By.XPATH,
            'css_selector': By.CSS_SELECTOR,
            'link_text'   : By.LINK_TEXT,
        }.get(mode)
        try:
            ele_list = WebDriverWait(self._driver, timeout).until(
                EC.presence_of_all_elements_located((tgt_mode, word))
            )
            return ele_list
        except Exception as e:
            return None
       
    # フレーム変更
    def switch_frame(self, iframe_ele):
        try:
            self._driver.switch_to.frame(iframe_ele)
            return self.close_driver
        except Exception as e:
            return None
    
    # セレクター取得
    @classmethod
    def get_select(self, ele):
        return Select(ele)
