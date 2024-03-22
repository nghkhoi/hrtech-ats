import re
import requests
import pykakasi
from time import sleep
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
import gspread
from oauth2client.service_account import ServiceAccountCredentials

COMPANY_NAME = ''
DRIVER_PATH = '/opt/chromedriver'
TRANSCRIPTION_COLUMN = '転記済み'

SERVICEACCOUNT_JSON = './file.json'
GSHEET_ID = ''

CHATWORK_ROOM_ID = ''
CHATWORK_TOKEN = ''

parse_datetime = lambda date_str: datetime(*map(int, re.findall(r'\d+', date_str)))

def convert_kana_word(input_text):
    kks = pykakasi.kakasi()
    result = kks.convert(input_text)
    return ''.join(entry['kana'] for entry in result if 'kana' in entry and entry['orig'] != ' ')

def convert_kana_name(full_name):
    parts = full_name.split(' ', 1)
    results = [convert_kana_word(part) for part in parts]
    return ' '.join(results)

# チャットワークメッセージ送信
def send_system_alert(msg, head_msg=True):
    session = requests.session()
    url = f'https://api.chatwork.com/v2/rooms/{CHATWORK_ROOM_ID}/messages'
    head_msg = '転記システムにエラーが発生しました！\n' if head_msg else ''
    body = f'[toall]\n{head_msg}' \
        f'発生関数：{COMPANY_NAME}\n' \
        f'{msg}\n\n'
    payload = {
        'body'       : body,
        'self_unread': 1,
    }
    headers = {
        'X-ChatWorkToken': CHATWORK_TOKEN,
        'Content-Type'   : 'application/x-www-form-urlencoded',
        'method'         : 'POST',
    }
    return session.post(url, data=payload, headers=headers)

# ユーザ情報を整形する
def get_user_info(worksheet, i, record, headers):
    # ユーザ情報の辞書
    userinfo = {}

    # エラーメッセージ格納用
    err_msg_list = []
    
    # 値取得用
    get_cell_data = lambda x: (data if (data := record[headers.index(x)]) not in ['#VALUE!', '#N/A'] else '') if x in headers else ''

    # お名前
    userinfo['name'] = get_cell_data('お名前')
    if not userinfo.get('name'): err_msg_list.append('お名前の値がありません')

    # フリガナ
    # userinfo['furigana'] = get_cell_data('フリガナ')
    # if not userinfo.get('furigana'):
    userinfo['furigana'] = convert_kana_name(userinfo['name'])
    if not userinfo.get('furigana'):
        err_msg_list.append('フリガナの値がありません')

    # 生年月日
    userinfo['birthday'] = get_cell_data('生年月日')
    if not userinfo.get('birthday'): err_msg_list.append('生年月日の値がありません')

    # メールアドレス
    userinfo['mail'] = get_cell_data('メールアドレス')
    if not userinfo.get('mail'): err_msg_list.append('メールアドレスの値がありません')

    # 電話番号
    userinfo['tel'] = get_cell_data('電話番号')
    if not userinfo.get('tel'): err_msg_list.append('電話番号の値がありません')

    # 転記URL
    userinfo['app_url'] = get_cell_data('転記URL')
    if not userinfo.get('app_url'): err_msg_list.append('転記URLの値がありません')

    # 必須項目が正常に取得できていなければエラーメッセージを送信
    if err_msg_list:
        err_msg_list.insert(0, f'{worksheet.title} {i}行目: パラメータエラー(下記項目は必須です)')
        err_msg_list.append(worksheet.url)
        err_msg = '\n'.join(err_msg_list)
        print(err_msg)
        send_system_alert(err_msg)
        idx = headers.index(TRANSCRIPTION_COLUMN) + 1
        worksheet.update_cell(i , idx, '転記エラー')
        return None
       
    return userinfo

def over_try_catch(f):
    def _wrapper(*args, **keywords):
        try:
            result = f(*args, **keywords)
            return result
        except Exception as e:
            err_msg = f'想定されていないエラーが発生しましたメッセージを確認してください\n{e}'
            if "has no attribute 'worksheet'" not in f'{e}':
                send_system_alert(err_msg)
            print(err_msg)
            return f'{e}'
    return _wrapper

@over_try_catch
def main():
    # グーグル認証オブジェクト取得
    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive',
    ]
    credentials = ServiceAccountCredentials.from_json_keyfile_name(SERVICEACCOUNT_JSON, scope)

    gc = gspread.authorize(credentials)

    # スプレッドシートワークシート取得
    worksheet = gc.open_by_key(GSHEET_ID).worksheet('応募者リスト')

    # スプシの内容取得
    all_record = worksheet.get_all_values()
    # ヘッダー取得
    headers = [re.sub('^([^\n]+)\n*(<|\().*', '\\1', header, flags=re.DOTALL).strip() for header in all_record[0]]
    # 転記ステータスの位置取得
    result_index = headers.index(TRANSCRIPTION_COLUMN)
    # 転記ステータスが空白のもののインデックスを取得
    startdate = datetime (2024, 1, 23)
    transfer_dates_dict = {
        i: record
        for i, record in enumerate(all_record[1:], start=2)
        if record[0] != '' and record[result_index] == '' and parse_datetime(record[0]) >= startdate
        }

    # ユーザ情報をアウトプットの形に整形する
    userinfo_dict = {
        i: user_data
        for i, record in transfer_dates_dict.items()
        if (user_data := get_user_info(worksheet, i, record, headers))}
    print('入力するユーザ情報取得完了')
    if not userinfo_dict:
        print('転記するデータはありません')
        return False

    # デバッグのための分岐
    options_list = ['--no-sandbox', '--disable-gpu', '--disable-dev-shm-usage', '--disable-dev-tools']

    if __name__ == '__main__':
        # ローカルマシンで呼び出す場合
        driver_path = r'D:\chromedriver\chromedriver.exe'
        # Set up Chrome service
        service = Service(executable_path=driver_path)
    else:
        # AWSラムダで呼び出す場合
        service = Service(ChromeDriverManager().install())
        options_list.extend(['--headless', '--lang=ja-JP', '--single-process'])
        chrome_options.binary_location = '/opt/chrome/chrome'

    chrome_options = Options()
    for option in options_list:
        chrome_options.add_argument(option)

    driver = webdriver.Chrome(service=service, options=chrome_options)
    print('クロームドライバ取得OK')

    # コンテンツが読み込まれるまで待機
    def load_wait(timeout=20, sleep_time=0):
        try:
            sleep(sleep_time)
            WebDriverWait(driver, timeout).until(EC.presence_of_all_elements_located)
        except Exception as e:
            print('タイムアウトしました')
            print(e)

    # 画面サイズ変更
    driver.maximize_window()

    # 転記ステータスのインデックス
    result_index = result_index + 1
 
    # 応募フォーム
    for i, userinfo in userinfo_dict.items():
        # ページ遷移
        try:
            app_url = userinfo.get('app_url')
            driver.get(app_url)
            load_wait(sleep_time=2)
            print(driver.current_url)

        except Exception as e:
            err_msg = f'{worksheet.title} {i}行目: 応募ページへの遷移に失敗しました\n{e}\n{worksheet.url}'
            print(err_msg)
            send_system_alert(err_msg)
            worksheet.update_cell(i , result_index, '転記エラー')
            continue

        # 各種項目入力
        try:
             # お名前
            now_input = 'お名前'
            driver.find_element(By.ID, 'name').send_keys(userinfo.get('name'))
            print(f'{i}行目: {now_input}入力完了')
            load_wait(sleep_time=1)

            # フリガナ
            now_input = 'フリガナ'
            furigana_input = driver.find_element(By.ID, 'name-kana')
            furigana_input.clear()
            furigana_input.send_keys(userinfo.get('furigana'))
            print(f'{i}行目: {now_input}入力完了')
            load_wait(sleep_time=1)

            # 生年月日
            now_input = '生年月日'
            birthday_str = userinfo.get('birthday').split("/")
            birthyear = birthday_str [0]
            birthmonth = str(int(birthday_str[1]))
            birthday = str(int(birthday_str[2]))
            birthyear_select = Select(driver.find_element(By.ID, 'form-item-03'))
            birthyear_select.select_by_value(birthyear)
            birthmonth_select = Select(driver.find_element(By.ID, 'form-item-04'))
            birthmonth_select.select_by_value(birthmonth)
            birthday_select = Select(driver.find_element(By.ID, 'form-item-05'))
            birthday_select.select_by_value(birthday)
            print(f'{i}行目: {now_input}入力完了')

            # メールアドレス
            now_input = 'メールアドレス'
            driver.find_element(By.ID, 'email').send_keys(userinfo.get('mail'))
            print(f'{i}行目: {now_input}入力完了')

            # 電話番号
            now_input = '電話番号'
            driver.find_element(By.ID, 'tel').send_keys(userinfo.get('tel'))
            print(f'{i}行目: {now_input}入力完了')

            # サービス利用規約同意
            now_input = 'サービス利用規約同意'
            driver.find_element(By.ID, 'agree').click()
            print(f'{i}行目: {now_input}入力完了')
            load_wait(sleep_time=1)

            # 同意して進む
            now_input = '同意して進む'
            driver.find_element(By.CSS_SELECTOR, 'button[data-entry-text="同意して進む"]').click()
            print(f'{i}行目: {now_input}入力完了')

        except Exception as e:
            err_msg = f'{i}行目: 「{now_input}」の入力に失敗しました\n{e}\n{worksheet.url}'
            print(err_msg)
            send_system_alert(err_msg)
            worksheet.update_cell(i, result_index, '転記エラー')
            continue

        # 登録処理
        try:
            load_wait(sleep_time=5)
            if 'PO_RegistCertificationA' not in driver.current_url:
                error = driver.find_element(By.CSS_SELECTOR, "div._error").text.replace("\n", " ").strip()
                worksheet.update_cell(i, result_index + 1, error)
                raise Exception('登録完了画面に遷移できませんでした')
            print(f'{i}行目: 処理済み')
            worksheet.update_cell(i, result_index, '転記済み')

        except Exception as e:
            err_msg = f'{i}行目: 処理失敗\n{type(e).__name__}\n{worksheet.url}'
            print(err_msg)
            send_system_alert(err_msg)
            worksheet.update_cell(i, result_index, '転記エラー')
            continue

def lambda_handler(event, lambda_context):
    main()

if __name__ == '__main__':
    lambda_handler({}, '')