WScript.Echo "VarType eXg"
WScript.Echo "vbEmpty     =" & vbEmpty     ' vbEmpty         0 Empty l (¢ú»)                                  
WScript.Echo "vbNull      =" & vbNull      ' vbNull          1 Null l (³øÈl)                                   
WScript.Echo "vbInteger   =" & vbInteger   ' vbInteger       2 ®^                                               
WScript.Echo "vbLong      =" & vbLong      ' vbLong          3 ·®^ (Long)                                      
WScript.Echo "vbSingle    =" & vbSingle    ' vbSingle        4 P¸x®¬_^ (Single)                        
WScript.Echo "vbDouble    =" & vbDouble    ' vbDouble        5 {¸x®¬_^ (Double)                        
WScript.Echo "vbCurrency  =" & vbCurrency  ' vbCurrency      6 ÊÝ^ (Currency)                                    
WScript.Echo "vbDate      =" & vbDate      ' vbDate          7 út^ (Date)                                        
WScript.Echo "vbString    =" & vbString    ' vbString        8 ¶ñ^                                             
WScript.Echo "vbObject    =" & vbObject    ' vbObject        9 I[g[V IuWFNg                        
WScript.Echo "vbError     =" & vbError     ' vbError        10 G[^                                             
WScript.Echo "vbBoolean   =" & vbBoolean   ' vbBoolean      11 u[^ (Boolean)                                   
WScript.Echo "vbVariant   =" & vbVariant   ' vbVariant      12 oAg^ (Variant) (oAg^zñÉÌÝgp)  
WScript.Echo "vbDataObject=" & vbDataObject' vbDataObject   13 ñI[g[V IuWFNg                      
WScript.Echo "vbByte      =" & vbByte      ' vbByte         17 oCg^                                             
WScript.Echo "vbArray     =" & vbArray     ' vbArray      8192 zñ (Array)                                         
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
