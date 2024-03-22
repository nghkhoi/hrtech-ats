import csv
import gspread
import re
from oauth2client.service_account import ServiceAccountCredentials
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

class GoogleServiceUtil:
    # コンストラクタ
    def __init__(self, json_file_path):
        self._credentials = self.create_credentials(json_file_path)
        self._drive = self.create_google_drive()
        self._gc = self.create_gspread_client()
    
    # 認証オブジェクト取得
    @property
    def credentials(self):
        return self._credentials
    
    # スプレッドシートのクライアント取得
    @property
    def gspread_client(self):
        return self._gc
    
    # グーグルドライブのコントローラ取得
    @property
    def drive(self):
        return self._drive

    # 認証情報オブジェクト生成
    def create_credentials(self, json_file_path):
        try:
            scope = [
                'https://spreadsheets.google.com/feeds',
                'https://www.googleapis.com/auth/drive',
            ]
            # サービスアカウント認証情報
            self._credentials = ServiceAccountCredentials.from_json_keyfile_name(json_file_path, scope)
        except Exception as e:
            print(f'認証情報オブジェクトの生成に失敗しました:{e}')
            self._credentials = None
        finally:
            return self._credentials

    # グーグルドライブのコントローラ生成
    def create_google_drive(self):
        try:
            # pydrive用の認証
            gauth = GoogleAuth()
            gauth.credentials = self._credentials
            self._drive = GoogleDrive(gauth)
        except Exception as e:
            print(f'グーグルドライブコントローラの生成に失敗しました:{e}')
            self._drive = None
        finally:
            return self._drive

    # スプレッドシートのクライアント生成
    def create_gspread_client(self):
        try:
            self._gc = gspread.authorize(self._credentials)
        except Exception as e:
            print(f'スプレッドシートのコントローラの生成に失敗しました:{e}')
            self._gc = None
        finally:
            return self._gc

    # グーグルドライブにスプレッドシート作成
    def add_spread_sheet(self, dir_id, title):
        try:
            # スプレッドシート作成
            f = self._drive.CreateFile({
                'title'   : title,
                'mimeType': 'application/vnd.google-apps.spreadsheet',
                'parents' : [{'id': dir_id}],
            })
            f.Upload()
            return f
        except Exception as e:
            print(f'スプレッドシートの作成に失敗しました:{e}')
            print(e)
            return None

    # グーグルドライブにディレクトリ作成
    def add_folder(self, dir_id, title):
        try:
            f = self._drive.CreateFile({
                'title'   : title,
                'mimeType': 'application/vnd.google-apps.folder',
                'parents' : [{'id': dir_id}],
            })
            f.Upload()
            return f
        except Exception as e:
            print(f'ディレクトリの作成に失敗しました:{e}')
            print(e)
            return None

    # スプレッドシートワークブック取得
    def get_spled_sheet_workbook(self, sheet_id):
        try:
            return self._gc.open_by_key(sheet_id)
        except Exception as e:
            print(f'スプレッドシートの取得に失敗しました:{e}')
            return None

    # リスト→スプレッドシート変換
    @classmethod
    def list_2_spred(self, add_body, worksheet, batch_size=1000, clear_mode=True):
        try:
            if clear_mode: worksheet.clear()
            # batch_size行ごとにデータを追加する
            for i in range(0, len(add_body), batch_size):
                range_label = f'A{i + 1}'
                subset = add_body[i: i + batch_size]
                worksheet.append_rows(subset, table_range=range_label)
            return worksheet
        except Exception as e:
            print(f'スプレッドシートの書き込みに失敗しました:{e}')
            return None

    # CSV→スプレッドシート変換
    @classmethod
    def csv_2_spred(self, csv_path, worksheet, batch_size=1000, encoding='utf_8', clear_mode=True):
        add_body = list(csv.reader(open(csv_path, encoding=encoding)))
        return self.list_2_spred(add_body, worksheet, batch_size, clear_mode)
    
    # DF→スプレッドシート変換
    @classmethod
    def df_2_spred(self, df, worksheet, batch_size=1000, header=True, clear_mode=True):
        add_body = df.values.tolist()
        if header: add_body.insert(0, df.columns.tolist())
        return self.list_2_spred(add_body, worksheet, batch_size, clear_mode)
   
    # ワークシート削除
    @classmethod
    def del_worksheet(self, workbook, sheet_name):
        try:
            worksheet = workbook.worksheet(sheet_name)
            if worksheet: workbook.del_worksheet(worksheet)
            return True
        except:
            return False
    
    # ワークシートに書き込み
    @classmethod
    def update_sheet(self, worksheet, data, start_cell='A1', batch_size=1000): 
        split_start_cell = re.sub("^([^0-9]+)([1-9]+)", "\\1,\\2", start_cell).split(",")
        if len(split_start_cell) != 2: return None
        try:
            # 開始位置を取得
            start_col, start_row = split_start_cell
            # 終了カラムを取得
            start_col_num = self.col_letter_to_num(start_col)
            end_col_num = start_col_num + len(data[0]) -1
            end_col = self.num_to_col_letter(end_col_num)
            data_list = [data[i: i+batch_size] for i in range(0, len(data), batch_size)]
            start_row = int(start_row)
            for split_data in data_list:
                end_row = start_row + batch_size - 1
                cell_range = f'{start_col}{start_row}:{end_col}{end_row}'
                # データをセルに書き込む
                worksheet.update(cell_range, split_data)
                start_row = end_row + 1
            return worksheet
        except Exception as e:
            return None

    # 列番号を文字列に変換
    @classmethod
    def num_to_col_letter(self, num):
        result = ""
        while num > 0:
            num, remainder = divmod(num - 1, 26)
            result = chr(65 + remainder) + result
        return result
    
    # 列名を数値に変換
    @classmethod
    def col_letter_to_num(self, col_letter):
        result = 0
        for i, char in enumerate(reversed(col_letter.upper())):
            result += (ord(char) - 64) * (26 ** i)
        return result
        
    # スプシのセルサイズ変更
    @classmethod
    def change_cell_size(self, worksheet, mode='ROWS', start_index=0, end_index=100, pixcel_size=20):
        mode = mode.upper()
        if mode != 'COLUMNS': mode='ROWS'
        try:
            requests = [{
                'updateDimensionProperties': {
                    'range': {
                        'sheetId'   : worksheet.id,
                        'dimension' : mode,
                        'startIndex': start_index, # 行の開始位置
                        'endIndex'  : end_index, # 行の終了位置
                    },
                    'properties': {
                        'pixelSize': pixcel_size, # セルの高さ（ピクセル単位）
                    },
                    'fields': 'pixelSize',
                }
            }]
            response = worksheet.spreadsheet.batch_update({'requests': requests})
            return worksheet
        except Exception as e:
            print(f'セルの縦幅変更に失敗しました\n{e}')
            return None
    
    # セルの色を変更
    @classmethod
    def change_cell_color(self, worksheet, color=[255,255,255], start_row=0, end_row=1, start_col=0, end_col=1):
        try:
            # 色取得
            red, green, blue = [rgb / 255 for rgb in color]
            # セルのバックグラウンドカラーを変更するリクエストを作成
            requests = [{
                'repeatCell': {
                    'range': {
                        'sheetId'         : worksheet.id,
                        'startRowIndex'   : start_row,
                        'endRowIndex'     : end_row,
                        'startColumnIndex': start_col,
                        'endColumnIndex'  : end_col,
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'backgroundColor': {
                                'red'  : red,
                                'green': green,
                                'blue' : blue,
                            }
                        }
                    },
                    'fields': 'userEnteredFormat.backgroundColor',
                }
            }]
            # リクエストを実行
            worksheet.spreadsheet.batch_update({'requests': requests})
            return worksheet
        except Exception as e:
            print(f'セルの色変更に失敗しました\n{e}')
            return None
