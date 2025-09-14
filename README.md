\# 🎵 wota-translater

快速剪輯御宅藝用副歌檔案的批次工具。  

支援 Excel 曲目清單、多段進入退出、自動淡入淡出、Peak 正規化。

====================

\## ✨ 功能特色

\- 依 Excel 清單批次裁切 `songs/` 音檔 → `output/` 320kbps MP3

\- 支援多段 `\[進入/退出]`，自動裁剪並拼接

\- 0:00 例外處理：若開頭與進入均為 0:00 → 不做淡入淡出

\- `to\_ms()` 支援單一秒數或 `mm:ss`

\- Peak 正規化至 -1.0 dBFS

\- 輸出錯誤紀錄（duplicate\_titles.txt, processing\_errors.txt ...）

====================

\## 📦 安裝與使用

1\. 直接下載 \*\*Release\*\* 的 `wota-translater.exe`（免安裝 Python）

&nbsp;	
	--在跟主程式資料夾下建立 songs資料夾 與 tracklist.xlsx。



&nbsp;	--將 songs/ 填充需要剪輯的音檔，並將 tracklist.xlsx 編輯好格式



&nbsp;	--運行wota-translater.exe



2\. 開發者可直接使用 `wota-translater.py`並更改參數進行微調


Excel 格式(tracklist.xlsx)

A(歌名)　B (開頭進入)　C (開頭退出)　D (第一段進入)　E (第一段退出)　F(第x段進入)……

歌名　　mm:ss　　　　mm:ss　　　　mm:ss　　　　mm:ss　　　　　mm:ss        

(注意!須確保Excel儲存格格式皆為[文字]，否則會處理失敗)


必填：
歌名(A)

可選：
開頭進入(B，選填)、開頭退出(C)
第一段進入(D)、第一段退出(E)
第X段進入、第X段退出（GHIJK……可無限擴充，[X]支援中文或阿拉伯數字）

====================

\## 📑輸出內容



`output/\*.mp3`
副歌成品


`duplicate\_titles.txt`
重複歌名

`duplicate\_matches.txt`
音檔被多歌名配對

`invalid\_time.txt`
時間格式 / 邏輯錯誤

`processing\_errors.txt`
音檔讀取 / 輸出異常

`unmatched\_titles.txt`
歌名找不到音檔

`unmatched\_audio.txt`
songs/ 中未被使用之檔案


剪輯後的音檔會存放在 `output/`，格式為 320 kbps MP3。


====================


\## ✂️剪輯邏輯：



只保留 \[進入 ~ 退出] 區段。



每段自動套用淡入/淡出。



0:00 例外：若首段 進入=0 → 不做淡入。



最後進行 Peak 正規化 −1 dBFS。













