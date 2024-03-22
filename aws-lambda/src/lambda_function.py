import re
from dateutil import parser
from datetime import date
from selenium.webdriver.common.by import By
from configparser import ConfigParser
from util.get_userinfo import *
from util.chrome_util import Chrome_Util
from util.google_service_util import GoogleServiceUtil

@over_try_catch
def main():
    # コンフィグファイル読み込み
    config = ConfigParser()
    config.read('./config.ini', encoding='utf-8')

    # グーグル認証JSON
    service_account_json = config.get('GOOGLE', 'service_account_json')
    # スプレッドシートID
    spread_sheet_id = config.get('GOOGLE', 'spread_sheet_id')
    # spread_sheet_id = config.get('GOOGLE', 'test_spread_sheet_id')
    # Webドライバパス
    driver_path = config.get('LOCAL', 'driver_path')
    # 転記ステータスカラム名
    transcription_name = config.get('LOCAL', 'transcription_name')
    print('コンフィグ読込完了')

    # グーグル認証オブジェクト取得
    gs_util = GoogleServiceUtil(service_account_json)

    # スプレッドシートワークシート取得
    worksheet = gs_util.get_spled_sheet_workbook(spread_sheet_id).worksheet('応募者リスト')
    # スプシの内容取得
    all_record = worksheet.get_all_values()
    # ヘッダー取得
    headers = [re.sub('^([^\n]+)\n*(<|\().*', '\\1', header, flags=re.DOTALL).strip() for header in all_record[0]]
    # 転記ステータスの位置取得
    result_index = headers.index(transcription_name)
    # 転記ステータスが空白のもののインデックスを取得
    today = date.today()
    transfer_dates_dict = {
        i: record
        for i, record in enumerate(all_record[1:], start=2)
        if record[0] != '' and record[result_index] == '' and parser.parse(record[0]).date() >= today
        }

    # ユーザ情報をアウトプットの形に整形する
    userinfo_dict = {
        i: user_data
        for i, record in transfer_dates_dict.items()
        if (user_data := get_user_info(config, worksheet, i, record, headers))}
    print('入力するユーザ情報取得完了')
    if not userinfo_dict:
        print('転記するデータはありません')
        return False

    # デバッグのための分岐
    options_list = [
        '--headless',
        '--lang=ja-JP',
        '--no-sandbox',
        '--disable-gpu',
        '--single-process',
        '--disable-dev-shm-usage',
        '--disable-dev-tools',
    ]
    chrome = Chrome_Util(driver_path, options_list, '/opt/chrome/chrome')
    print('クロームドライバ取得OK')

    # 画面サイズ変更
    chrome.driver.maximize_window()

    # 転記ステータスのインデックス
    result_index = result_index + 1
 
    # 応募フォーム
    for i, userinfo in userinfo_dict.items():
        # ページ遷移
        try:
            app_url = userinfo.get('app_url')
            chrome.location_href(app_url)
            chrome.load_wait(sleep_time=2)
            print(chrome.current_url)

        except Exception as e:
            err_msg = f'{worksheet.title} {i}行目: 応募ページへの遷移に失敗しました\n{e}\n{worksheet.url}'
            print(err_msg)
            send_system_alert(config, err_msg)
            worksheet.update_cell(i , result_index, '転記エラー')
            continue

        # 各種項目入力
        try:
             # お名前
            now_input = 'お名前'
            chrome.find_element('id', 'name').send_keys(userinfo.get('name'))
            print(f'{i}行目: {now_input}入力完了')
            chrome.load_wait(sleep_time=1)

            # フリガナ
            now_input = 'フリガナ'
            furigana_input = chrome.find_element('id', 'name-kana')
            furigana_input.clear()
            furigana_input.send_keys(userinfo.get('furigana'))
            print(f'{i}行目: {now_input}入力完了')
            chrome.load_wait(sleep_time=1)

            # 生年月日
            now_input = '生年月日'
            birthday = parser.parse(userinfo.get('birthday'))
            birthyear_select = chrome.get_select(chrome.find_element('id', 'form-item-03'))
            birthyear_select.select_by_value(str(birthday.year))
            birthmonth_select = chrome.get_select(chrome.find_element('id', 'form-item-04'))
            birthmonth_select.select_by_value(str(birthday.month))
            birthday_select = chrome.get_select(chrome.find_element('id', 'form-item-05'))
            birthday_select.select_by_value(str(birthday.day))
            print(f'{i}行目: {now_input}入力完了')

            # メールアドレス
            now_input = 'メールアドレス'
            chrome.find_element('id', 'email').send_keys(userinfo.get('mail'))
            print(f'{i}行目: {now_input}入力完了')

            # 電話番号
            now_input = '電話番号'
            chrome.find_element('id', 'tel').send_keys(userinfo.get('tel'))
            print(f'{i}行目: {now_input}入力完了')

            # サービス利用規約同意
            now_input = 'サービス利用規約同意'
            chrome.find_element('id', 'agree').click()
            print(f'{i}行目: {now_input}入力完了')
            chrome.load_wait(sleep_time=1)

            # 同意して進む
            now_input = '同意して進む'
            chrome.find_element('css_selector', 'button[data-entry-text="同意して進む"]').click()
            print(f'{i}行目: {now_input}入力完了')

        except Exception as e:
            err_msg = f'{i}行目: 「{now_input}」の入力に失敗しました\n{e}\n{worksheet.url}'
            print(err_msg)
            send_system_alert(config, err_msg)
            worksheet.update_cell(i, result_index, '転記エラー')
            continue
        
        # 登録処理
        try:
            chrome.load_wait(sleep_time=2)
            success_xpath = "//p[contains(text(), 'まだ登録は完了していません。')]"
            if chrome.find_element('title', success_xpath) is None:
                raise Exception('登録完了画面に遷移できませんでした')
            print(f'{i}行目: 処理済み')
            worksheet.update_cell(i, result_index, '転記済み')

        except Exception as e:
            err_msg = f'{i}行目: 処理失敗\n{e}\n{worksheet.url}'
            print(err_msg)
            send_system_alert(config, err_msg)
            worksheet.update_cell(i, result_index, '転記エラー')
            continue

def lambda_handler(event, lambda_context):
    main()

if __name__ == '__main__':
    lambda_handler({}, '')