# Yuanta-API 元大期貨交易API串接與自動化交易

## !!! 注意檔案目錄安裝路徑要正確，否則容易出現報錯

程式檔案存放位置(C:/Yuanta/QAPI)

vs-code 終端機預設目錄(C:/Users/User)下方存放wxPython.whl檔案

anaconda 要看自己安裝的位置，去設定python-32bits環境

======
- 開戶階段-選擇身分別為保險公司-可以拿到行情API權限
- !一般用戶(自營商交易帳戶)-無法取得權限，只能透過API下單與看盤軟體[例如:元大點金靈](https://www.yuanta.com.tw/eYuanta/securities/aporder/Instructions/836878aa-5e5f-4dc8-9d18-984e9bf5c1cd?TargetId=16b7b99c-7cc5-4b05-9208-58f026f8da0a&TargetMode=2)下單
- !交易api (http://easywin.yuantafutures.com.tw/download/YuantaOrdAPI_py.zip)

## 環境設定:
- python3.9
- anaconda 64-bits
- create python 32-bits environment
  - 開啟anaconda prompt, 輸入以下指令創建新環境
- 詳情可看(https://www.youtube.com/watch?v=o1Zzw-Y_n2g)
```
set CONDA_FORCE_32BIT=1 #名稱自定
set #開始建置
conda create -n ENV_32BIT #建立環境
activate ENV_32BIT #啟動新環境，繼續在環境中安裝
conda install python=3.9 #於新環境中設置python3.9，由於api中套件wx,comtypes支援3.9以下版本
```
  - 可以打開C:\anaconda\envs\ENV_32BIT\python.exe，成功顯示python=3.9
- open anaconda, change environment to ENV_32BIT 
![image](https://user-images.githubusercontent.com/72643996/231789700-59a87c0f-c7e1-449a-b7bb-651635b08afb.png)
- download wxpython 32-bits
(https://files.pythonhosted.org/packages/05/d1/40ca8bacba94d49d3c379a0ea0220cb9594f45978a40c321a46d4727d93e/wxPython-4.1.1-cp39-cp39-win32.whl)
- install comtypes (do it in anaonda prompt) 
```
conda install -c free comtypes
```
- from anaconda launch jupyter or spyder or vs-code, running Yuantaapi.py
- after install wxpython 32-bits, move whl file to vs-code terminal directory (C:\Users\User)
- open Yuantaapi.py through anaconda in vs-code and type in vs-code terminal
```
pip install wxPython-4.1.1-cp39-cp39-win32.whl
```
- 執行後跳出登入視窗如下圖所示
![image](https://user-images.githubusercontent.com/72643996/228880830-90aeef64-2446-4c3a-a591-924e091ec88a.png)
- 輸入元大期貨帳戶與交易密碼
- 右欄註冊區輸入產品代碼[元大期貨下單商品代碼規則](https://www.multicharts.com.tw/dis/dis_Content.aspx?D_ID=3&SN=5120)
- 成功連線與註冊
- ![image](https://user-images.githubusercontent.com/72643996/231790557-2ac98109-b0a8-4b9c-9e48-65ad0cef36b2.png)
- 商品註冊成功
- ![image](https://user-images.githubusercontent.com/72643996/230544057-a19ea185-4888-4796-b1cd-51a78d609ada.png)
