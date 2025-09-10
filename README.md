# 台灣測速照相設備資料比較系統

## 專案概述

本專案是一個自動化的台灣測速照相設備資料收集、處理和比較系統。系統能夠從各縣市政府網站自動下載測速照相設備資料，進行資料清理和標準化處理，並與第三方資料（如 Mio 導航資料）進行比較分析。

## 主要功能

### 1. 資料收集 (Selenium)
- 自動化從全台 22 個縣市政府的官方網站下載測速照相設備資料
- 支援 PDF 和 Excel 格式的檔案下載
- 使用 Selenium WebDriver 進行網頁自動化操作
- **新增功能**: 完整的錯誤處理機制，單一縣市失敗不影響其他縣市處理
- **新增功能**: Debug 模式，可單獨測試特定縣市

### 2. 資料轉換 (Converter)
- 將 PDF 檔案中的表格資料轉換為 Excel 格式
- 處理不同格式的表格結構
- 自動識別和清理資料格式

### 3. 資料處理 (Split Data)
- 使用正則表達式解析地址資訊
- 將地址分解為：縣市、區域、道路1、道路2、備註等欄位
- 支援 18 種不同的地址格式解析規則
- 去除重複資料和無效資料

### 4. 資料比較 (Compare Data)
- 比較政府官方資料與第三方資料（如 Mio）
- 使用模糊匹配演算法進行地址比對
- 標記匹配結果和差異
- 產生比較報告

## 專案結構

```
TWN_speedcam_compare/
├── selenium/                    # 資料收集模組
│   ├── selenium_main.py        # 主要執行檔案
│   ├── city_processor.py       # 各縣市資料處理器
│   ├── city_py/               # 各縣市專用處理腳本
│   └── city_data/             # 下載的原始資料
├── converter/                  # 資料轉換模組
│   ├── pdf_to_excel.py        # PDF 轉 Excel 工具
│   ├── process_taipei_data.py # 台北市資料特殊處理
│   └── city_data/             # 轉換後的 Excel 資料
├── split_data/                 # 資料處理模組
│   ├── split_tool.py          # 主要處理工具
│   ├── step_split.py          # 地址解析規則
│   ├── data_process*.py       # 資料處理步驟
│   ├── mio_data_process.py    # Mio 資料處理
│   └── standard_in/           # 標準化後的資料
└── compare_data/               # 資料比較模組
    ├── compare_main.py        # 比較主程式
    ├── data/                  # 輸入資料
    ├── input/                 # 處理後的資料
    └── output/                # 比較結果
```

## 支援的縣市

系統支援以下 22 個縣市的資料收集：

- 台北市 (taipei)
- 新北市 (newTaipei)
- 桃園市 (TaoYuan)
- 新竹市 (HsinChu_web)
- 新竹縣 (HsinChu)
- 苗栗縣 (MiaoLi)
- 台中市 (TaiChung)
- 彰化縣 (ChangHua)
- 南投縣 (NanTou)
- 雲林縣 (YunLin)
- 嘉義市 (ChiaYi_sh)
- 嘉義縣 (ChiaYi)
- 台南市 (TaiNan)
- 高雄市 (Kaohsiung)
- 屏東縣 (PingTung)
- 宜蘭縣 (YiLan)
- 花蓮縣 (HuaLian)
- 台東縣 (TaiTung)
- 基隆市 (KeeLung)
- 新竹科學園區 (Science_Park)
- 金門縣 (KinMen)
- 澎湖縣 (PengHu)

## 安裝需求

### Python 套件
```bash
pip install pandas
pip install openpyxl
pip install selenium
pip install pdfplumber
pip install fuzzywuzzy
pip install python-dateutil
```

### 系統需求
- Python 3.7+
- Chrome 瀏覽器
- ChromeDriver（需與 Chrome 版本匹配）

## 使用方式

### 1. 資料收集

#### 正常模式（處理所有縣市）
```bash
cd selenium
python selenium_main.py
```

#### Debug 模式（只處理特定縣市）
```bash
cd selenium
# 只處理桃園市
python selenium_main.py TaoYuan

# 只處理台北市
python selenium_main.py taipei

# 只處理高雄市
python selenium_main.py Kaohsiung
```

**可用的縣市名稱**: taipei, TaoYuan, KeeLung, Science_Park, HsinChu_web, MiaoLi, TaiChung, YunLin, newTaipei, HsinChu, ChiaYi, NanTou, TaiNan, TaiTung, HuaLian, YiLan, KinMen, Kaohsiung, ChiaYi_sh, PengHu, ChangHua

### 2. 資料轉換
```bash
cd converter
python pdf_to_excel.py
```

### 3. 資料處理
```bash
cd split_data
python split_tool.py
```

### 4. 資料比較
```bash
cd compare_data
python compare_main.py
```

## 資料格式

### 輸入資料格式
系統處理的資料包含以下欄位：
- `city`: 縣市
- `district`: 區域
- `road_1`: 主要道路
- `road_2`: 次要道路
- `note`: 備註資訊
- `camera_type`: 照相機類型（1=固定式，5=移動式）
- `speed`: 限速
- `address`: 完整地址

### 輸出資料格式
比較結果包含：
- 原始資料欄位
- 匹配狀態 (`match`: 0=未匹配，1=部分匹配，2=完全匹配)
- 政府資料對應欄位 (`g_city`, `g_district`, `g_road_1`, `g_road_2`, `g_note`)
- 差異標記（紅色=照相機類型不同，黃色=限速不同）

## 技術特色

### 1. 智能地址解析
- 使用 18 種正則表達式規則解析不同格式的地址
- 支援繁體中文地址的各種表達方式
- 自動處理地址中的特殊符號和格式

### 2. 模糊匹配演算法
- 使用 `fuzzywuzzy` 進行地址相似度比對
- 設定不同的相似度閾值（道路1: 90%，道路2: 85%）
- 智能處理備註資訊的匹配

### 3. 自動化資料清理
- 去除重複資料
- 標準化 Unicode 字元
- 清理無效和異常資料

### 4. 視覺化結果
- 使用顏色標記差異（紅色、黃色）
- 產生易讀的 Excel 報告
- 支援多種輸出格式

### 5. 錯誤處理與穩定性
- **智能錯誤處理**: 單一縣市處理失敗時，程式會記錄錯誤並繼續處理其他縣市
- **詳細錯誤報告**: 提供完整的處理結果摘要，包括成功和失敗的縣市清單
- **資源管理**: 確保 WebDriver 正確關閉，避免資源洩漏
- **超時控制**: 設定頁面載入超時時間，避免程式卡住

## 注意事項

1. **網路連線**: 資料收集需要穩定的網路連線
2. **Chrome 版本**: 確保 ChromeDriver 與 Chrome 瀏覽器版本匹配
3. **資料更新**: 政府網站結構可能變更，需要定期更新爬蟲腳本
4. **法律合規**: 請遵守各網站的使用條款和爬蟲政策
5. **Debug 模式**: 開發和測試時建議使用 Debug 模式，只處理特定縣市以提高效率
6. **錯誤處理**: 程式具備完整的錯誤處理機制，即使部分縣市失敗也會繼續處理其他縣市

## 新功能說明

### Debug 模式
Debug 模式讓開發者可以單獨測試特定縣市，提高開發和除錯效率：

```bash
# 測試桃園市
python selenium_main.py TaoYuan

# 測試台北市
python selenium_main.py taipei
```

### 錯誤處理機制
- **繼續執行**: 即使某個縣市處理失敗，程式會繼續處理其他縣市
- **詳細記錄**: 記錄每個縣市的處理狀態和錯誤訊息
- **結果摘要**: 在程式結束時提供完整的處理結果摘要

### 處理結果範例
```
==================================================
資料下載處理完成摘要
==================================================
✅ 成功處理的縣市 (20 個):
   - taipei
   - TaoYuan
   - KeeLung
   ...

❌ 處理失敗的縣市 (2 個):
   - MiaoLi: 頁面載入超時
   - TaiNan: 找不到網頁元素

總計: 20 成功, 2 失敗
==================================================
```

## 維護與更新

- 定期檢查各縣市網站結構變更
- 更新地址解析規則以支援新的地址格式
- 優化匹配演算法以提高準確率
- 新增支援的縣市和資料來源
- 持續改善錯誤處理機制和穩定性

## 授權

本專案僅供學習和研究使用，請遵守相關法律法規和網站使用條款。

## 聯絡資訊

如有問題或建議，請聯繫專案維護者。
