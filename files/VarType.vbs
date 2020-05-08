WScript.Echo "VarType テスト"
WScript.Echo "vbEmpty     =" & vbEmpty     '│vbEmpty     │   0│Empty 値 (未初期化)                                 │
WScript.Echo "vbNull      =" & vbNull      '│vbNull      │   1│Null 値 (無効な値)                                  │
WScript.Echo "vbInteger   =" & vbInteger   '│vbInteger   │   2│整数型                                              │
WScript.Echo "vbLong      =" & vbLong      '│vbLong      │   3│長整数型 (Long)                                     │
WScript.Echo "vbSingle    =" & vbSingle    '│vbSingle    │   4│単精度浮動小数点数型 (Single)                       │
WScript.Echo "vbDouble    =" & vbDouble    '│vbDouble    │   5│倍精度浮動小数点数型 (Double)                       │
WScript.Echo "vbCurrency  =" & vbCurrency  '│vbCurrency  │   6│通貨型 (Currency)                                   │
WScript.Echo "vbDate      =" & vbDate      '│vbDate      │   7│日付型 (Date)                                       │
WScript.Echo "vbString    =" & vbString    '│vbString    │   8│文字列型                                            │
WScript.Echo "vbObject    =" & vbObject    '│vbObject    │   9│オートメーション オブジェクト                       │
WScript.Echo "vbError     =" & vbError     '│vbError     │  10│エラー型                                            │
WScript.Echo "vbBoolean   =" & vbBoolean   '│vbBoolean   │  11│ブール型 (Boolean)                                  │
WScript.Echo "vbVariant   =" & vbVariant   '│vbVariant   │  12│バリアント型 (Variant) (バリアント型配列にのみ使用) │
WScript.Echo "vbDataObject=" & vbDataObject'│vbDataObject│  13│非オートメーション オブジェクト                     │
WScript.Echo "vbByte      =" & vbByte      '│vbByte      │  17│バイト型                                            │
WScript.Echo "vbArray     =" & vbArray     '│vbArray     │8192│配列 (Array)                                        │
dim	v
WScript.Echo "VarType(v)=" & VarType(v)
v = Null
WScript.Echo "VarType(v)=" & VarType(v)
v = 123
WScript.Echo "VarType(v)=" & VarType(v)
v = 123456789
WScript.Echo "VarType(v)=" & VarType(v)
v = 1.1
WScript.Echo "VarType(v)=" & VarType(v)
