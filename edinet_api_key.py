import os

def get_edinet_api_key():
    """
    .edinet_api_key_config から EDINET API キーを取得します。
    ファイルが存在しない場合や空の場合は環境変数を参照します。
    """
    # スクリプトのディレクトリ
    base_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(base_dir, '.edinet_api_key_config')
    
    # 1. ファイルから取得
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                key = f.read().strip()
                if key:
                    return key
        except Exception:
            pass
            
    # 2. 環境変数から取得
    return os.environ.get('EDINET_API_KEY', '')

# 直接インポートするための変数定義
EDINET_API_KEY = get_edinet_api_key()
