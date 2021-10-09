#檢查項目
inspect = [ '外觀','運動黏度(40°C)','運動黏度(100°C)',
'水分','閃火點(開口)' ,'總酸值' ,'機械雜質']

#輸入欄位的值['檢驗方法' ,'物性指標' ,'檢驗結果' ,'單位']
presentVal_1 = ['目測', '微黃色油狀透明液體' , '微黃色油狀透明液體','-' ,'' ] #'外觀'
presentVal_2 = ['ASTM D445', '≤23' , '17.71','cSt' ,'' ] #'運動黏度(40°C)'
presentVal_3 = ['ASTM D445', '≤4.2' , '3.43','cSt' ,'' ] #'運動黏度(100°C)'
presentVal_4 = ['ASTM D6304', '≤300' , '122.81','mg/L' ,'' ] #'水分'
presentVal_5 = ['ASTM D92', '≧180' , '202','℃' ,'' ] #'開火點(開口)'
presentVal_6 = ['ASTM D664', '≤1.0' , '0.0002','mgKOH/g' ,'' ] #'總酸值'
presentVal_7 = ['ASTM D473', '≤0.03' , '0.0100','%' ,'' ] #'機械雜質'

#修改成二維陣列
presentVal = [['目測', '微黃色油狀透明液體' , '微黃色油狀透明液體','-' ,'' ],
              ['ASTM D445', '≤23' , '17.71','cSt' ,'' ],
              ['ASTM D445', '≤4.2' , '3.43','cSt' ,'' ],
              ['ASTM D6304', '≤300' , '122.81','mg/L' ,'' ],
              ['ASTM D92', '≧180' , '202','℃' ,'' ],
              ['ASTM D664', '≤1.0' , '0.0002','mgKOH/g' ,'' ],
              ['ASTM D473', '≤0.03' , '0.0100','%' ,'' ]]


#標題欄位
col_name = ['檢驗項目','檢驗方法' ,'物性指標' ,'檢驗結果' ,'單位'
,'檢驗人']