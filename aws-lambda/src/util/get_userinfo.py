import re
import requests
import pykakasi

def convert_kana_word(input_text):
    kks = pykakasi.kakasi()
    result = kks.convert(input_text)
    return ''.join(entry['kana'] for entry in result if 'kana' in entry and entry['orig'] != ' ')

def convert_kana_name(full_name):
    parts = full_name.split(' ', 1)
    results = [convert_kana_word(part) for part in parts]
    return ' '.join(results)

# 大枠のエラー通知用デコレータ
def over_try_catch(f):
    def _wrapper(*args, **keywords):
        try:
            result = f(*args, **keywords)
            return result
        except Exception as e:
            # コンフィグファイル読み込み
            from configparser import ConfigParser
            config = ConfigParser()
            config.read('./config.ini', encoding='utf-8')
            err_msg = f'想定されていないエラーが発生しましたメッセージを確認してください\n{e}'
            if "has no attribute 'worksheet'" not in f'{e}':
                send_system_alert(config, err_msg)
            print(err_msg)
            return f'{e}'
    return _wrapper


# チャットワークメッセージ送信
def send_system_alert(config, msg, head_msg=True):
    company_name = config.get('LOCAL', 'company_name')
    room_id = config.get('CHAT_WORK', 'room_id')
    token = config.get('CHAT_WORK', 'token')
    session = requests.session()
    url = f'https://api.chatwork.com/v2/rooms/{room_id}/messages'
    head_msg = '転記システムにエラーが発生しました！\n' if head_msg else ''
    body = f'[toall]\n{head_msg}' \
        f'発生関数：{company_name}\n' \
        f'{msg}\n\n'
    payload = {
        'body'       : body,
        'self_unread': 1,
    }
    headers = {
        'X-ChatWorkToken': token,
        'Content-Type'   : 'application/x-www-form-urlencoded',
        'method'         : 'POST',
    }
    return session.post(url, data=payload, headers=headers)


# ユーザ情報を整形する
def get_user_info(config, worksheet, i, record, headers):
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
        send_system_alert(config, err_msg)
        idx = headers.index(config.get('LOCAL', 'transcription_name')) + 1
        worksheet.update_cell(i , idx, '転記エラー')
        return None
       
    return userinfo
