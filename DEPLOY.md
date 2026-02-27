# サーバー構築手順 (makoto.xtomo.com)

このディレクトリを `makoto.xtomo.com` サーバーに配置し、Webアプリケーションとして公開する手順です。

## 1. 必要な環境のインストール
サーバー上でPython環境（venv）を作成し、必要なライブラリをインストールします。

```bash
# プロジェクトディレクトリに移動
cd /path/to/work_office

# 仮想環境の作成と有効化
python3 -m venv venv
source venv/bin/activate

# 必須ライブラリのインストール
pip install -r requirements.txt
```

## 2. アプリケーションの仮起動テスト
設定に問題がないか、手動で起動してみます。
```bash
python app.py
```
エラーなく起動したら `Ctrl+C` で終了します。

## 3. 本番環境での運用 (Gunicorn)
Flask内蔵サーバーは開発用のため、本番向けには `gunicorn` を使用します。
```bash
pip install gunicorn

# Gunicornでアプリをデーモン化して起動（例: 8000ポート、ワーカー4つ）
gunicorn -w 4 -b 127.0.0.1:8000 app:app --daemon
```

## 4. (任意) Nginx / Apache との連携
外部（学生）からアクセスできるようにするには、NginxやApacheのリバースプロキシ設定で、ドメインへのリクエストを先ほど立ち上げたポート `8000` に転送（ProxyPass）します。

### Nginx の設定例
```nginx
server {
    listen 443 ssl;
    server_name makoto.xtomo.com;

    # SSL証明書等の設定群 ...

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
}
```

### Apache の設定例
```apache
<VirtualHost *:443>
    ServerName makoto.xtomo.com
    
    # SSL証明書等の設定群 ...

    ProxyPreserveHost On
    ProxyPass / http://127.0.0.1:8000/
    ProxyPassReverse / http://127.0.0.1:8000/
</VirtualHost>
```

以上の設定で、学生が `https://makoto.xtomo.com/` にアクセスし、ZIPファイルをアップロードすると自動でExcelが変換・ダウンロードされるようになります。

---

# コアサーバー（V2）への配置手順

コアサーバーには標準で `python3.10` がインストールされているため、これを利用して環境構築を行います。

## 1. サーバー上での準備 (venv)

SSHでサーバーにログインし、以下の順にコマンドを実行してください。

```bash
# 1. プロジェクトディレクトリの作成と移動
mkdir -p ~/public_html/xbrl2excel
cd ~/public_html/xbrl2excel

# 2. 仮想環境の作成 (システムのpython3.10を使用)
python3.10 -m venv venv

# 3. 仮想環境を有効化してライブラリをインストール
source venv/bin/activate
pip install -r requirements.txt
```

## 2. ファイルの転送とパーミッション

以下のファイルを `public_html/xbrl2excel` に配置してください。
- `app.py`
- `convert_xbrl_to_excel.py`
- `index.cgi`
- `.htaccess`
- `requirements.txt`
- `templates/` (ディレクトリごと)
- `edinet_taxonomies/` (ディレクトリごと)

転送後、CGIの実行権限を付与します：
```bash
chmod 755 index.cgi
```

## 3. ファイルの修正

### index.cgi
1行目のパスをサーバーの実際のユーザー名と設置パスに合わせて書き換えてください。
```python
#!/home/（あなたのユーザー名）/public_html/xbrl2excel/venv/bin/python3
```

## 4. 動作確認

ブラウザで `https://makoto.xtomo.com/` にアクセスし、画面が表示されるか確認してください。
動かない場合は、SSHで以下を実行してエラーを確認します：
```bash
cd ~/public_html/xbrl2excel
./index.cgi
```

### .htaccess
リライト設定が有効であることを確認してください。

## 4. 動作確認

ブラウザで `https://makoto.xtomo.com/` にアクセスし、画面が表示されるか確認してください。
動かない場合は、SSHで以下を実行してエラーを確認します：
```bash
cd ~/public_html/xbrl2excel
./index.cgi
```
