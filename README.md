# VALORANT-Sensitivity-Tracker
～データと統計で、エイムの「正解」を見つける～

・プロジェクト概要 / Project Overview

このアプリは、統計的なパフォーマンスデータを用いて、プレイヤーが自分に最適なマウス感度を科学的に特定するための「試合後レビュー」ツールです。

・主な機能 / Key Features

感度の自動記録 (Smart OCR): ホットキーを押すだけで、ゲーム内設定画面から感度を自動抽出。

スタッツ同期 (Match History): HenrikDev APIを使用し、直近の試合データを自動取得。

安定度分析 (Stability Analysis): 平均ACSと標準偏差から、あなたのエイムの「ブレ」を数値化。

プロ・ベンチマーク (Benchmarks): Meiy選手など、トッププロの統計値と比較可能。

・導入方法 / How to Use

ダウンロード: 右側の「Releases」から最新の .exe ファイルをダウンロードしてください。

APIキーの設定: https://github.com/Henrik-3/unofficial-valorant-api?tab=readme-ov-file　からDiscordに入室しAPIキーを入手。
アプリ内の設定タブに、自身の HenrikDev API Key を入力します。
APIキーの取得手順については、こちらの解説ブログ(https://scrub.blog/valorant-rank-overlay-20240923/)が非常に分かりやすいです。

Tesseract OCRの導入: 文字認識を行うため、PCに Tesseract OCR がインストールされている必要があります。
※ .exe単体で動作しますが、文字読み取り機能にのみ必要です。
1,以下のURLからインストーラーをダウンロード
https://github.com/UB-Mannheim/tesseract/wiki
2,tesseract-ocr-w64-setup-5.x.x.exe を実行
（設定は全部デフォルトのままNextを押すだけでOK）
3,「Tesseractパス」の欄にインストール先のパスを入力して保存
パスの確認方法：
デフォルトでインストールした場合は以下
C:\Program Files\Tesseract-OCR\tesseract.exe
わからない場合はエクスプローラーで tesseract.exe を検索して、そのフルパスを入力してください

※ .exe単体で動作しますが、文字読み取り機能にのみ必要です。
