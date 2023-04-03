# Yuanta-API 元大期貨交易API串接與自動化交易

- 開戶階段-選擇身分別為保險公司-可以拿到行情API權限
- !一般用戶(自營商交易帳戶)-無法取得權限，只能透過API下單與看盤軟體[例如:元大點金靈](https://www.yuanta.com.tw/eYuanta/securities/aporder/Instructions/836878aa-5e5f-4dc8-9d18-984e9bf5c1cd?TargetId=16b7b99c-7cc5-4b05-9208-58f026f8da0a&TargetMode=2)下單

![image](https://user-images.githubusercontent.com/72643996/228870955-dfd3e9a6-9d13-4e08-b760-d4520212307f.png)

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
- download wxpython 32-bits
(https://files.pythonhosted.org/packages/05/d1/40ca8bacba94d49d3c379a0ea0220cb9594f45978a40c321a46d4727d93e/wxPython-4.1.1-cp39-cp39-win32.whl)
- install comtypes (do it in anaonda prompt : conda install -c free comtypes)
- activigate jupyter or spyder or vs-code, for running Yuantaapi.py
- after install wxpython 32-bits, move whl file to vs-code terminal directory (C:\Users\User)
- open Yuantaapi.py through anaconda in vs-code and type in vs-code terminal
```
pip install wxPython-4.1.1-cp39-cp39-win32.whl
```
- 執行後跳出登入視窗如下圖所示
![image](https://user-images.githubusercontent.com/72643996/228880830-90aeef64-2446-4c3a-a591-924e091ec88a.png)
- 輸入元大期貨帳戶與交易密碼
- 右欄註冊區輸入產品代碼(e.g. 小型台指FIMTX)
- 成功連線與註冊
- ![api連線成功](https://user-images.githubusercontent.com/72643996/229060823-ae01c109-6e4a-4b47-8660-a81cc4144802.png)

