import webview
import datetime
import random
import threading
import pypyodbc

def read_kebiao(db_name, password, sql):
    # import pypyodbc
    # Driver={Microsoft Access Driver (*.mdb, *.accdb)};PWD=

    try:
        _str = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)};PWD=' + password + ";DBQ=" + db_name
        conn = pypyodbc.win_connect_mdb(_str)
        cur = conn.cursor()
        cur.execute(sql)
        dbdata = cur.fetchall()
        cur.commit()
        cur.close()  # 关闭游标
        conn.close()
        # print(dbdata)
        return dbdata
    except pypyodbc.Error as e:
        # print('查询课表出错！' + e.__str__())
        print('查询课表出错！' + e.__str__())
        return []

def return_random_num():
    # import datetime
    # import random
    _nowTime = datetime.datetime.now().strftime("%Y%m%d%H%M%S")  # 生成当前时间
    _randomNum = random.randint(0, 100)  # 生成的随机整数n，其中0<=n<=100
    if _randomNum <= 10:
        _randomNum = str(0) + str(_randomNum)
    _uniqueNum = str(_nowTime) + str(_randomNum)
    return _uniqueNum

def loadhtml(backcolor='', td_list=''):
    table_template = '''<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/
    xhtml1-transitional.dtd">
    <html xmlns="http://www.w3.org/1999/xhtml">
    <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>标准课表</title>
    <style type="text/css">
    th {
        font-family: "宋体";
    }
    </style>
    <script>
    function fill_td() {
        var td_list=#td_list;
        //获得每周的上课节点(必须给每天的每个节次td设定id)，改色
        //id=week(1-25)_day(1-7)_jieci(1-5):
        //12_3_5：第12周周三晚上那个td格子
        var text="";
        for (var i=1;i<=td_list.length;i++) 
            {var room_tdArr = document.getElementById(td_list[i-1]);
            if(room_tdArr.style.backgroundColor != "mediumseagreen")
                {room_tdArr.style.backgroundColor = "mediumseagreen";}
            else 
                {text=text + td_list[i-1] +",";}
            }
        //if(text.length>0){alert(text+ "课程时间冲突"); }      
        }
    
    
    function showResponse(response) {
        //alert(response.message);
        var name='W';
        var weeksLength=response.message.length
        //alert('本学期共有'+String(weeksLength)+'周');
        for (var i=1;i<=25;i++)
            {name='W'+i.toString();
             var tdArr = document.getElementById(name);
             //表格左边第一列改底色
            tdArr.innerText.substr(0,1) == "D" ?tdArr.style.backgroundColor ="#back__color" : tdArr.style.backgroundColor = "white";
            if(i<=weeksLength)
                //填写周次日期区间到表格左边第一列
                {tdArr.innerText=response.message[i-1];}
            else
                {tdArr.innerText='';}
            }
        }
    function getWeeks(){
        pywebview.api.getWeeks().then(showResponse);
        fill_td();
        }
</script>
    </head>

    <body onload="getWeeks()">
    <table  border="1" cellspacing="0" cellpadding="0">
      <tr bgcolor="seashell">
        <th width="220" rowspan="2" scope="col">日期</th>
        <th width="20" rowspan="2" align="center" scope="col">周序</th>
        <th colspan="5" scope="col">一</th>
        <th colspan="5" scope="col">二</th>
        <th colspan="5" scope="col">三</th>
        <th colspan="5" scope="col">四</th>
        <th colspan="5" scope="col">五</th>
        <th colspan="5" scope="col">六</th>
        <th colspan="5" scope="col">日</th>
      </tr>
      <tr bgcolor="seashell">
        <th width="20" align="center" valign="middle" scope="col">12</th>
        <th width="20" align="center" valign="middle" scope="col">34</th>
        <th width="20" align="center" valign="middle" scope="col">56</th>
        <th width="20" align="center" valign="middle" scope="col">78</th>
        <th width="20" align="center" valign="middle" scope="col">晚</th>
        <th width="20" align="center" valign="middle" scope="col">12</th>
        <th width="20" align="center" valign="middle" scope="col">34</th>
        <th width="20" align="center" valign="middle" scope="col">56</th>
        <th width="20" align="center" valign="middle" scope="col">78</th>
        <th width="20" align="center" valign="middle" scope="col">晚</th>
        <th width="20" align="center" valign="middle" scope="col">12</th>
        <th width="20" align="center" valign="middle" scope="col">34</th>
        <th width="20" align="center" valign="middle" scope="col">56</th>
        <th width="20" align="center" valign="middle" scope="col">78</th>
        <th width="20" align="center" valign="middle" scope="col">晚</th>
        <th width="20" align="center" valign="middle" scope="col">12</th>
        <th width="20" align="center" valign="middle" scope="col">34</th>
        <th width="20" align="center" valign="middle" scope="col">56</th>
        <th width="20" align="center" valign="middle" scope="col">78</th>
        <th width="20" align="center" valign="middle" scope="col">晚</th>
        <th width="20" align="center" valign="middle" scope="col">12</th>
        <th width="20" align="center" valign="middle" scope="col">34</th>
        <th width="20" align="center" valign="middle" scope="col">56</th>
        <th width="20" align="center" valign="middle" scope="col">78</th>
        <th width="20" align="center" valign="middle" scope="col">晚</th>
        <th width="20" align="center" valign="middle" scope="col">12</th>
        <th width="20" align="center" valign="middle" scope="col">34</th>
        <th width="20" align="center" valign="middle" scope="col">56</th>
        <th width="20" align="center" valign="middle" scope="col">78</th>
        <th width="20" align="center" valign="middle" scope="col">晚</th>
        <th width="20" align="center" valign="middle" scope="col">12</th>
        <th width="20" align="center" valign="middle" scope="col">34</th>
        <th width="20" align="center" valign="middle" scope="col">56</th>
        <th width="20" align="center" valign="middle" scope="col">78</th>
        <th width="20" align="center" valign="middle" scope="col">晚</th>
      </tr>
      <tr>
        <td width="220" align="center" id="W1">D1</td>
        <td width="20" align="center">1</td>
        <td id="1_1_1" width="20">&nbsp;</td>
        <td id="1_1_2" width="20">&nbsp;</td>
        <td id="1_1_3" width="20">&nbsp;</td>
        <td id="1_1_4" width="20">&nbsp;</td>
        <td id="1_1_5" width="20">&nbsp;</td>
        <td id="1_2_1">&nbsp;</td>
        <td id="1_2_2">&nbsp;</td>
        <td id="1_2_3">&nbsp;</td>
        <td id="1_2_4">&nbsp;</td>
        <td id="1_2_5">&nbsp;</td>
        <td id="1_3_1">&nbsp;</td>
        <td id="1_3_2">&nbsp;</td>
        <td id="1_3_3">&nbsp;</td>
        <td id="1_3_4">&nbsp;</td>
        <td id="1_3_5">&nbsp;</td>
        <td id="1_4_1">&nbsp;</td>
        <td id="1_4_2">&nbsp;</td>
        <td id="1_4_3">&nbsp;</td>
        <td id="1_4_4">&nbsp;</td>
        <td id="1_4_5">&nbsp;</td>
        <td id="1_5_1">&nbsp;</td>
        <td id="1_5_2">&nbsp;</td>
        <td id="1_5_3">&nbsp;</td>
        <td id="1_5_4">&nbsp;</td>
        <td id="1_5_5">&nbsp;</td>
        <td id="1_6_1">&nbsp;</td>
        <td id="1_6_2">&nbsp;</td>
        <td id="1_6_3">&nbsp;</td>
        <td id="1_6_4">&nbsp;</td>
        <td id="1_6_5">&nbsp;</td>
        <td id="1_7_1">&nbsp;</td>
        <td id="1_7_2">&nbsp;</td>
        <td id="1_7_3">&nbsp;</td>
        <td id="1_7_4">&nbsp;</td>
        <td id="1_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center" id="W2">D2</td>
        <td width="20" align="center">2</td>
        <td id="2_1_1" width="20">&nbsp;</td>
        <td id="2_1_2" width="20">&nbsp;</td>
        <td id="2_1_3" width="20">&nbsp;</td>
        <td id="2_1_4" width="20">&nbsp;</td>
        <td id="2_1_5" width="20">&nbsp;</td>
        <td id="2_2_1">&nbsp;</td>
        <td id="2_2_2">&nbsp;</td>
        <td id="2_2_3">&nbsp;</td>
        <td id="2_2_4">&nbsp;</td>
        <td id="2_2_5">&nbsp;</td>
        <td id="2_3_1">&nbsp;</td>
        <td id="2_3_2">&nbsp;</td>
        <td id="2_3_3">&nbsp;</td>
        <td id="2_3_4">&nbsp;</td>
        <td id="2_3_5">&nbsp;</td>
        <td id="2_4_1">&nbsp;</td>
        <td id="2_4_2">&nbsp;</td>
        <td id="2_4_3">&nbsp;</td>
        <td id="2_4_4">&nbsp;</td>
        <td id="2_4_5">&nbsp;</td>
        <td id="2_5_1">&nbsp;</td>
        <td id="2_5_2">&nbsp;</td>
        <td id="2_5_3">&nbsp;</td>
        <td id="2_5_4">&nbsp;</td>
        <td id="2_5_5">&nbsp;</td>
        <td id="2_6_1">&nbsp;</td>
        <td id="2_6_2">&nbsp;</td>
        <td id="2_6_3">&nbsp;</td>
        <td id="2_6_4">&nbsp;</td>
        <td id="2_6_5">&nbsp;</td>
        <td id="2_7_1">&nbsp;</td>
        <td id="2_7_2">&nbsp;</td>
        <td id="2_7_3">&nbsp;</td>
        <td id="2_7_4">&nbsp;</td>
        <td id="2_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center" id="W3">D3</td>
        <td width="20" align="center">3</td>
        <td id="3_1_1" width="20">&nbsp;</td>
        <td id="3_1_2" width="20">&nbsp;</td>
        <td id="3_1_3" width="20">&nbsp;</td>
        <td id="3_1_4" width="20">&nbsp;</td>
        <td id="3_1_5" width="20">&nbsp;</td>
        <td id="3_2_1">&nbsp;</td>
        <td id="3_2_2">&nbsp;</td>
        <td id="3_2_3">&nbsp;</td>
        <td id="3_2_4">&nbsp;</td>
        <td id="3_2_5">&nbsp;</td>
        <td id="3_3_1">&nbsp;</td>
        <td id="3_3_2">&nbsp;</td>
        <td id="3_3_3">&nbsp;</td>
        <td id="3_3_4">&nbsp;</td>
        <td id="3_3_5">&nbsp;</td>
        <td id="3_4_1">&nbsp;</td>
        <td id="3_4_2">&nbsp;</td>
        <td id="3_4_3">&nbsp;</td>
        <td id="3_4_4">&nbsp;</td>
        <td id="3_4_5">&nbsp;</td>
        <td id="3_5_1">&nbsp;</td>
        <td id="3_5_2">&nbsp;</td>
        <td id="3_5_3">&nbsp;</td>
        <td id="3_5_4">&nbsp;</td>
        <td id="3_5_5">&nbsp;</td>
        <td id="3_6_1">&nbsp;</td>
        <td id="3_6_2">&nbsp;</td>
        <td id="3_6_3">&nbsp;</td>
        <td id="3_6_4">&nbsp;</td>
        <td id="3_6_5">&nbsp;</td>
        <td id="3_7_1">&nbsp;</td>
        <td id="3_7_2">&nbsp;</td>
        <td id="3_7_3">&nbsp;</td>
        <td id="3_7_4">&nbsp;</td>
        <td id="3_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center"  id="W4">D4</td>
        <td width="20" align="center">4</td>
        <td id="4_1_1" width="20">&nbsp;</td>
        <td id="4_1_2" width="20">&nbsp;</td>
        <td id="4_1_3" width="20">&nbsp;</td>
        <td id="4_1_4" width="20">&nbsp;</td>
        <td id="4_1_5" width="20">&nbsp;</td>
        <td id="4_2_1">&nbsp;</td>
        <td id="4_2_2">&nbsp;</td>
        <td id="4_2_3">&nbsp;</td>
        <td id="4_2_4">&nbsp;</td>
        <td id="4_2_5">&nbsp;</td>
        <td id="4_3_1">&nbsp;</td>
        <td id="4_3_2">&nbsp;</td>
        <td id="4_3_3">&nbsp;</td>
        <td id="4_3_4">&nbsp;</td>
        <td id="4_3_5">&nbsp;</td>
        <td id="4_4_1">&nbsp;</td>
        <td id="4_4_2">&nbsp;</td>
        <td id="4_4_3">&nbsp;</td>
        <td id="4_4_4">&nbsp;</td>
        <td id="4_4_5">&nbsp;</td>
        <td id="4_5_1">&nbsp;</td>
        <td id="4_5_2">&nbsp;</td>
        <td id="4_5_3">&nbsp;</td>
        <td id="4_5_4">&nbsp;</td>
        <td id="4_5_5">&nbsp;</td>
        <td id="4_6_1">&nbsp;</td>
        <td id="4_6_2">&nbsp;</td>
        <td id="4_6_3">&nbsp;</td>
        <td id="4_6_4">&nbsp;</td>
        <td id="4_6_5">&nbsp;</td>
        <td id="4_7_1">&nbsp;</td>
        <td id="4_7_2">&nbsp;</td>
        <td id="4_7_3">&nbsp;</td>
        <td id="4_7_4">&nbsp;</td>
        <td id="4_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center" id="W5">D5</td>
        <td width="20" align="center">5</td>
        <td id="5_1_1" width="20">&nbsp;</td>
        <td id="5_1_2" width="20">&nbsp;</td>
        <td id="5_1_3" width="20">&nbsp;</td>
        <td id="5_1_4" width="20">&nbsp;</td>
        <td id="5_1_5" width="20">&nbsp;</td>
        <td id="5_2_1">&nbsp;</td>
        <td id="5_2_2">&nbsp;</td>
        <td id="5_2_3">&nbsp;</td>
        <td id="5_2_4">&nbsp;</td>
        <td id="5_2_5">&nbsp;</td>
        <td id="5_3_1">&nbsp;</td>
        <td id="5_3_2">&nbsp;</td>
        <td id="5_3_3">&nbsp;</td>
        <td id="5_3_4">&nbsp;</td>
        <td id="5_3_5">&nbsp;</td>
        <td id="5_4_1">&nbsp;</td>
        <td id="5_4_2">&nbsp;</td>
        <td id="5_4_3">&nbsp;</td>
        <td id="5_4_4">&nbsp;</td>
        <td id="5_4_5">&nbsp;</td>
        <td id="5_5_1">&nbsp;</td>
        <td id="5_5_2">&nbsp;</td>
        <td id="5_5_3">&nbsp;</td>
        <td id="5_5_4">&nbsp;</td>
        <td id="5_5_5">&nbsp;</td>
        <td id="5_6_1">&nbsp;</td>
        <td id="5_6_2">&nbsp;</td>
        <td id="5_6_3">&nbsp;</td>
        <td id="5_6_4">&nbsp;</td>
        <td id="5_6_5">&nbsp;</td>
        <td id="5_7_1">&nbsp;</td>
        <td id="5_7_2">&nbsp;</td>
        <td id="5_7_3">&nbsp;</td>
        <td id="5_7_4">&nbsp;</td>
        <td id="5_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center" id="W6">D6</td>
        <td width="20" align="center">6</td>
        <td id="6_1_1" width="20">&nbsp;</td>
        <td id="6_1_2" width="20">&nbsp;</td>
        <td id="6_1_3" width="20">&nbsp;</td>
        <td id="6_1_4" width="20">&nbsp;</td>
        <td id="6_1_5" width="20">&nbsp;</td>
        <td id="6_2_1">&nbsp;</td>
        <td id="6_2_2">&nbsp;</td>
        <td id="6_2_3">&nbsp;</td>
        <td id="6_2_4">&nbsp;</td>
        <td id="6_2_5">&nbsp;</td>
        <td id="6_3_1">&nbsp;</td>
        <td id="6_3_2">&nbsp;</td>
        <td id="6_3_3">&nbsp;</td>
        <td id="6_3_4">&nbsp;</td>
        <td id="6_3_5">&nbsp;</td>
        <td id="6_4_1">&nbsp;</td>
        <td id="6_4_2">&nbsp;</td>
        <td id="6_4_3">&nbsp;</td>
        <td id="6_4_4">&nbsp;</td>
        <td id="6_4_5">&nbsp;</td>
        <td id="6_5_1">&nbsp;</td>
        <td id="6_5_2">&nbsp;</td>
        <td id="6_5_3">&nbsp;</td>
        <td id="6_5_4">&nbsp;</td>
        <td id="6_5_5">&nbsp;</td>
        <td id="6_6_1">&nbsp;</td>
        <td id="6_6_2">&nbsp;</td>
        <td id="6_6_3">&nbsp;</td>
        <td id="6_6_4">&nbsp;</td>
        <td id="6_6_5">&nbsp;</td>
        <td id="6_7_1">&nbsp;</td>
        <td id="6_7_2">&nbsp;</td>
        <td id="6_7_3">&nbsp;</td>
        <td id="6_7_4">&nbsp;</td>
        <td id="6_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center" id="W7">D7</td>
        <td width="20" align="center">7</td>
        <td id="7_1_1" width="20">&nbsp;</td>
        <td id="7_1_2" width="20">&nbsp;</td>
        <td id="7_1_3" width="20">&nbsp;</td>
        <td id="7_1_4" width="20">&nbsp;</td>
        <td id="7_1_5" width="20">&nbsp;</td>
        <td id="7_2_1">&nbsp;</td>
        <td id="7_2_2">&nbsp;</td>
        <td id="7_2_3">&nbsp;</td>
        <td id="7_2_4">&nbsp;</td>
        <td id="7_2_5">&nbsp;</td>
        <td id="7_3_1">&nbsp;</td>
        <td id="7_3_2">&nbsp;</td>
        <td id="7_3_3">&nbsp;</td>
        <td id="7_3_4">&nbsp;</td>
        <td id="7_3_5">&nbsp;</td>
        <td id="7_4_1">&nbsp;</td>
        <td id="7_4_2">&nbsp;</td>
        <td id="7_4_3">&nbsp;</td>
        <td id="7_4_4">&nbsp;</td>
        <td id="7_4_5">&nbsp;</td>
        <td id="7_5_1">&nbsp;</td>
        <td id="7_5_2">&nbsp;</td>
        <td id="7_5_3">&nbsp;</td>
        <td id="7_5_4">&nbsp;</td>
        <td id="7_5_5">&nbsp;</td>
        <td id="7_6_1">&nbsp;</td>
        <td id="7_6_2">&nbsp;</td>
        <td id="7_6_3">&nbsp;</td>
        <td id="7_6_4">&nbsp;</td>
        <td id="7_6_5">&nbsp;</td>
        <td id="7_7_1">&nbsp;</td>
        <td id="7_7_2">&nbsp;</td>
        <td id="7_7_3">&nbsp;</td>
        <td id="7_7_4">&nbsp;</td>
        <td id="7_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center" id="W8">D8</td>
        <td width="20" align="center">8</td>
        <td id="8_1_1" width="20">&nbsp;</td>
        <td id="8_1_2" width="20">&nbsp;</td>
        <td id="8_1_3" width="20">&nbsp;</td>
        <td id="8_1_4" width="20">&nbsp;</td>
        <td id="8_1_5" width="20">&nbsp;</td>
        <td id="8_2_1">&nbsp;</td>
        <td id="8_2_2">&nbsp;</td>
        <td id="8_2_3">&nbsp;</td>
        <td id="8_2_4">&nbsp;</td>
        <td id="8_2_5">&nbsp;</td>
        <td id="8_3_1">&nbsp;</td>
        <td id="8_3_2">&nbsp;</td>
        <td id="8_3_3">&nbsp;</td>
        <td id="8_3_4">&nbsp;</td>
        <td id="8_3_5">&nbsp;</td>
        <td id="8_4_1">&nbsp;</td>
        <td id="8_4_2">&nbsp;</td>
        <td id="8_4_3">&nbsp;</td>
        <td id="8_4_4">&nbsp;</td>
        <td id="8_4_5">&nbsp;</td>
        <td id="8_5_1">&nbsp;</td>
        <td id="8_5_2">&nbsp;</td>
        <td id="8_5_3">&nbsp;</td>
        <td id="8_5_4">&nbsp;</td>
        <td id="8_5_5">&nbsp;</td>
        <td id="8_6_1">&nbsp;</td>
        <td id="8_6_2">&nbsp;</td>
        <td id="8_6_3">&nbsp;</td>
        <td id="8_6_4">&nbsp;</td>
        <td id="8_6_5">&nbsp;</td>
        <td id="8_7_1">&nbsp;</td>
        <td id="8_7_2">&nbsp;</td>
        <td id="8_7_3">&nbsp;</td>
        <td id="8_7_4">&nbsp;</td>
        <td id="8_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center" id="W9">D9</td>
        <td width="20" align="center">9</td>
        <td id="9_1_1" width="20">&nbsp;</td>
        <td id="9_1_2" width="20">&nbsp;</td>
        <td id="9_1_3" width="20">&nbsp;</td>
        <td id="9_1_4" width="20">&nbsp;</td>
        <td id="9_1_5" width="20">&nbsp;</td>
        <td id="9_2_1">&nbsp;</td>
        <td id="9_2_2">&nbsp;</td>
        <td id="9_2_3">&nbsp;</td>
        <td id="9_2_4">&nbsp;</td>
        <td id="9_2_5">&nbsp;</td>
        <td id="9_3_1">&nbsp;</td>
        <td id="9_3_2">&nbsp;</td>
        <td id="9_3_3">&nbsp;</td>
        <td id="9_3_4">&nbsp;</td>
        <td id="9_3_5">&nbsp;</td>
        <td id="9_4_1">&nbsp;</td>
        <td id="9_4_2">&nbsp;</td>
        <td id="9_4_3">&nbsp;</td>
        <td id="9_4_4">&nbsp;</td>
        <td id="9_4_5">&nbsp;</td>
        <td id="9_5_1">&nbsp;</td>
        <td id="9_5_2">&nbsp;</td>
        <td id="9_5_3">&nbsp;</td>
        <td id="9_5_4">&nbsp;</td>
        <td id="9_5_5">&nbsp;</td>
        <td id="9_6_1">&nbsp;</td>
        <td id="9_6_2">&nbsp;</td>
        <td id="9_6_3">&nbsp;</td>
        <td id="9_6_4">&nbsp;</td>
        <td id="9_6_5">&nbsp;</td>
        <td id="9_7_1">&nbsp;</td>
        <td id="9_7_2">&nbsp;</td>
        <td id="9_7_3">&nbsp;</td>
        <td id="9_7_4">&nbsp;</td>
        <td id="9_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center" id="W10">D10</td>
        <td width="20" align="center">10</td>
        <td id="10_1_1" width="20">&nbsp;</td>
        <td id="10_1_2" width="20">&nbsp;</td>
        <td id="10_1_3" width="20">&nbsp;</td>
        <td id="10_1_4" width="20">&nbsp;</td>
        <td id="10_1_5" width="20">&nbsp;</td>
        <td id="10_2_1">&nbsp;</td>
        <td id="10_2_2">&nbsp;</td>
        <td id="10_2_3">&nbsp;</td>
        <td id="10_2_4">&nbsp;</td>
        <td id="10_2_5">&nbsp;</td>
        <td id="10_3_1">&nbsp;</td>
        <td id="10_3_2">&nbsp;</td>
        <td id="10_3_3">&nbsp;</td>
        <td id="10_3_4">&nbsp;</td>
        <td id="10_3_5">&nbsp;</td>
        <td id="10_4_1">&nbsp;</td>
        <td id="10_4_2">&nbsp;</td>
        <td id="10_4_3">&nbsp;</td>
        <td id="10_4_4">&nbsp;</td>
        <td id="10_4_5">&nbsp;</td>
        <td id="10_5_1">&nbsp;</td>
        <td id="10_5_2">&nbsp;</td>
        <td id="10_5_3">&nbsp;</td>
        <td id="10_5_4">&nbsp;</td>
        <td id="10_5_5">&nbsp;</td>
        <td id="10_6_1">&nbsp;</td>
        <td id="10_6_2">&nbsp;</td>
        <td id="10_6_3">&nbsp;</td>
        <td id="10_6_4">&nbsp;</td>
        <td id="10_6_5">&nbsp;</td>
        <td id="10_7_1">&nbsp;</td>
        <td id="10_7_2">&nbsp;</td>
        <td id="10_7_3">&nbsp;</td>
        <td id="10_7_4">&nbsp;</td>
        <td id="10_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center" id="W11">D11</td>
        <td width="20" align="center">11</td>
        <td id="11_1_1" width="20">&nbsp;</td>
        <td id="11_1_2" width="20">&nbsp;</td>
        <td id="11_1_3" width="20">&nbsp;</td>
        <td id="11_1_4" width="20">&nbsp;</td>
        <td id="11_1_5" width="20">&nbsp;</td>
        <td id="11_2_1">&nbsp;</td>
        <td id="11_2_2">&nbsp;</td>
        <td id="11_2_3">&nbsp;</td>
        <td id="11_2_4">&nbsp;</td>
        <td id="11_2_5">&nbsp;</td>
        <td id="11_3_1">&nbsp;</td>
        <td id="11_3_2">&nbsp;</td>
        <td id="11_3_3">&nbsp;</td>
        <td id="11_3_4">&nbsp;</td>
        <td id="11_3_5">&nbsp;</td>
        <td id="11_4_1">&nbsp;</td>
        <td id="11_4_2">&nbsp;</td>
        <td id="11_4_3">&nbsp;</td>
        <td id="11_4_4">&nbsp;</td>
        <td id="11_4_5">&nbsp;</td>
        <td id="11_5_1">&nbsp;</td>
        <td id="11_5_2">&nbsp;</td>
        <td id="11_5_3">&nbsp;</td>
        <td id="11_5_4">&nbsp;</td>
        <td id="11_5_5">&nbsp;</td>
        <td id="11_6_1">&nbsp;</td>
        <td id="11_6_2">&nbsp;</td>
        <td id="11_6_3">&nbsp;</td>
        <td id="11_6_4">&nbsp;</td>
        <td id="11_6_5">&nbsp;</td>
        <td id="11_7_1">&nbsp;</td>
        <td id="11_7_2">&nbsp;</td>
        <td id="11_7_3">&nbsp;</td>
        <td id="11_7_4">&nbsp;</td>
        <td id="11_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center" id="W12">D12</td>
        <td width="20" align="center">12</td>
        <td id="12_1_1" width="20">&nbsp;</td>
        <td id="12_1_2" width="20">&nbsp;</td>
        <td id="12_1_3" width="20">&nbsp;</td>
        <td id="12_1_4" width="20">&nbsp;</td>
        <td id="12_1_5" width="20">&nbsp;</td>
        <td id="12_2_1">&nbsp;</td>
        <td id="12_2_2">&nbsp;</td>
        <td id="12_2_3">&nbsp;</td>
        <td id="12_2_4">&nbsp;</td>
        <td id="12_2_5">&nbsp;</td>
        <td id="12_3_1">&nbsp;</td>
        <td id="12_3_2">&nbsp;</td>
        <td id="12_3_3">&nbsp;</td>
        <td id="12_3_4">&nbsp;</td>
        <td id="12_3_5">&nbsp;</td>
        <td id="12_4_1">&nbsp;</td>
        <td id="12_4_2">&nbsp;</td>
        <td id="12_4_3">&nbsp;</td>
        <td id="12_4_4">&nbsp;</td>
        <td id="12_4_5">&nbsp;</td>
        <td id="12_5_1">&nbsp;</td>
        <td id="12_5_2">&nbsp;</td>
        <td id="12_5_3">&nbsp;</td>
        <td id="12_5_4">&nbsp;</td>
        <td id="12_5_5">&nbsp;</td>
        <td id="12_6_1">&nbsp;</td>
        <td id="12_6_2">&nbsp;</td>
        <td id="12_6_3">&nbsp;</td>
        <td id="12_6_4">&nbsp;</td>
        <td id="12_6_5">&nbsp;</td>
        <td id="12_7_1">&nbsp;</td>
        <td id="12_7_2">&nbsp;</td>
        <td id="12_7_3">&nbsp;</td>
        <td id="12_7_4">&nbsp;</td>
        <td id="12_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center" id="W13">D13</td>
        <td width="20" align="center">13</td>
        <td id="13_1_1" width="20">&nbsp;</td>
        <td id="13_1_2" width="20">&nbsp;</td>
        <td id="13_1_3" width="20">&nbsp;</td>
        <td id="13_1_4" width="20">&nbsp;</td>
        <td id="13_1_5" width="20">&nbsp;</td>
        <td id="13_2_1">&nbsp;</td>
        <td id="13_2_2">&nbsp;</td>
        <td id="13_2_3">&nbsp;</td>
        <td id="13_2_4">&nbsp;</td>
        <td id="13_2_5">&nbsp;</td>
        <td id="13_3_1">&nbsp;</td>
        <td id="13_3_2">&nbsp;</td>
        <td id="13_3_3">&nbsp;</td>
        <td id="13_3_4">&nbsp;</td>
        <td id="13_3_5">&nbsp;</td>
        <td id="13_4_1">&nbsp;</td>
        <td id="13_4_2">&nbsp;</td>
        <td id="13_4_3">&nbsp;</td>
        <td id="13_4_4">&nbsp;</td>
        <td id="13_4_5">&nbsp;</td>
        <td id="13_5_1">&nbsp;</td>
        <td id="13_5_2">&nbsp;</td>
        <td id="13_5_3">&nbsp;</td>
        <td id="13_5_4">&nbsp;</td>
        <td id="13_5_5">&nbsp;</td>
        <td id="13_6_1">&nbsp;</td>
        <td id="13_6_2">&nbsp;</td>
        <td id="13_6_3">&nbsp;</td>
        <td id="13_6_4">&nbsp;</td>
        <td id="13_6_5">&nbsp;</td>
        <td id="13_7_1">&nbsp;</td>
        <td id="13_7_2">&nbsp;</td>
        <td id="13_7_3">&nbsp;</td>
        <td id="13_7_4">&nbsp;</td>
        <td id="13_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center" id="W14">D14</td>
        <td width="20" align="center">14</td>
        <td id="14_1_1" width="20">&nbsp;</td>
        <td id="14_1_2" width="20">&nbsp;</td>
        <td id="14_1_3" width="20">&nbsp;</td>
        <td id="14_1_4" width="20">&nbsp;</td>
        <td id="14_1_5" width="20">&nbsp;</td>
        <td id="14_2_1">&nbsp;</td>
        <td id="14_2_2">&nbsp;</td>
        <td id="14_2_3">&nbsp;</td>
        <td id="14_2_4">&nbsp;</td>
        <td id="14_2_5">&nbsp;</td>
        <td id="14_3_1">&nbsp;</td>
        <td id="14_3_2">&nbsp;</td>
        <td id="14_3_3">&nbsp;</td>
        <td id="14_3_4">&nbsp;</td>
        <td id="14_3_5">&nbsp;</td>
        <td id="14_4_1">&nbsp;</td>
        <td id="14_4_2">&nbsp;</td>
        <td id="14_4_3">&nbsp;</td>
        <td id="14_4_4">&nbsp;</td>
        <td id="14_4_5">&nbsp;</td>
        <td id="14_5_1">&nbsp;</td>
        <td id="14_5_2">&nbsp;</td>
        <td id="14_5_3">&nbsp;</td>
        <td id="14_5_4">&nbsp;</td>
        <td id="14_5_5">&nbsp;</td>
        <td id="14_6_1">&nbsp;</td>
        <td id="14_6_2">&nbsp;</td>
        <td id="14_6_3">&nbsp;</td>
        <td id="14_6_4">&nbsp;</td>
        <td id="14_6_5">&nbsp;</td>
        <td id="14_7_1">&nbsp;</td>
        <td id="14_7_2">&nbsp;</td>
        <td id="14_7_3">&nbsp;</td>
        <td id="14_7_4">&nbsp;</td>
        <td id="14_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center" id="W15">D15</td>
        <td width="20" align="center">15</td>
        <td id="15_1_1" width="20">&nbsp;</td>
        <td id="15_1_2" width="20">&nbsp;</td>
        <td id="15_1_3" width="20">&nbsp;</td>
        <td id="15_1_4" width="20">&nbsp;</td>
        <td id="15_1_5" width="20">&nbsp;</td>
        <td id="15_2_1">&nbsp;</td>
        <td id="15_2_2">&nbsp;</td>
        <td id="15_2_3">&nbsp;</td>
        <td id="15_2_4">&nbsp;</td>
        <td id="15_2_5">&nbsp;</td>
        <td id="15_3_1">&nbsp;</td>
        <td id="15_3_2">&nbsp;</td>
        <td id="15_3_3">&nbsp;</td>
        <td id="15_3_4">&nbsp;</td>
        <td id="15_3_5">&nbsp;</td>
        <td id="15_4_1">&nbsp;</td>
        <td id="15_4_2">&nbsp;</td>
        <td id="15_4_3">&nbsp;</td>
        <td id="15_4_4">&nbsp;</td>
        <td id="15_4_5">&nbsp;</td>
        <td id="15_5_1">&nbsp;</td>
        <td id="15_5_2">&nbsp;</td>
        <td id="15_5_3">&nbsp;</td>
        <td id="15_5_4">&nbsp;</td>
        <td id="15_5_5">&nbsp;</td>
        <td id="15_6_1">&nbsp;</td>
        <td id="15_6_2">&nbsp;</td>
        <td id="15_6_3">&nbsp;</td>
        <td id="15_6_4">&nbsp;</td>
        <td id="15_6_5">&nbsp;</td>
        <td id="15_7_1">&nbsp;</td>
        <td id="15_7_2">&nbsp;</td>
        <td id="15_7_3">&nbsp;</td>
        <td id="15_7_4">&nbsp;</td>
        <td id="15_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center" id="W16">D16</td>
        <td width="20" align="center">16</td>
        <td id="16_1_1" width="20">&nbsp;</td>
        <td id="16_1_2" width="20">&nbsp;</td>
        <td id="16_1_3" width="20">&nbsp;</td>
        <td id="16_1_4" width="20">&nbsp;</td>
        <td id="16_1_5" width="20">&nbsp;</td>
        <td id="16_2_1">&nbsp;</td>
        <td id="16_2_2">&nbsp;</td>
        <td id="16_2_3">&nbsp;</td>
        <td id="16_2_4">&nbsp;</td>
        <td id="16_2_5">&nbsp;</td>
        <td id="16_3_1">&nbsp;</td>
        <td id="16_3_2">&nbsp;</td>
        <td id="16_3_3">&nbsp;</td>
        <td id="16_3_4">&nbsp;</td>
        <td id="16_3_5">&nbsp;</td>
        <td id="16_4_1">&nbsp;</td>
        <td id="16_4_2">&nbsp;</td>
        <td id="16_4_3">&nbsp;</td>
        <td id="16_4_4">&nbsp;</td>
        <td id="16_4_5">&nbsp;</td>
        <td id="16_5_1">&nbsp;</td>
        <td id="16_5_2">&nbsp;</td>
        <td id="16_5_3">&nbsp;</td>
        <td id="16_5_4">&nbsp;</td>
        <td id="16_5_5">&nbsp;</td>
        <td id="16_6_1">&nbsp;</td>
        <td id="16_6_2">&nbsp;</td>
        <td id="16_6_3">&nbsp;</td>
        <td id="16_6_4">&nbsp;</td>
        <td id="16_6_5">&nbsp;</td>
        <td id="16_7_1">&nbsp;</td>
        <td id="16_7_2">&nbsp;</td>
        <td id="16_7_3">&nbsp;</td>
        <td id="16_7_4">&nbsp;</td>
        <td id="16_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center" id="W17">D17</td>
        <td width="20" align="center">17</td>
        <td id="17_1_1" width="20">&nbsp;</td>
        <td id="17_1_2" width="20">&nbsp;</td>
        <td id="17_1_3" width="20">&nbsp;</td>
        <td id="17_1_4" width="20">&nbsp;</td>
        <td id="17_1_5" width="20">&nbsp;</td>
        <td id="17_2_1">&nbsp;</td>
        <td id="17_2_2">&nbsp;</td>
        <td id="17_2_3">&nbsp;</td>
        <td id="17_2_4">&nbsp;</td>
        <td id="17_2_5">&nbsp;</td>
        <td id="17_3_1">&nbsp;</td>
        <td id="17_3_2">&nbsp;</td>
        <td id="17_3_3">&nbsp;</td>
        <td id="17_3_4">&nbsp;</td>
        <td id="17_3_5">&nbsp;</td>
        <td id="17_4_1">&nbsp;</td>
        <td id="17_4_2">&nbsp;</td>
        <td id="17_4_3">&nbsp;</td>
        <td id="17_4_4">&nbsp;</td>
        <td id="17_4_5">&nbsp;</td>
        <td id="17_5_1">&nbsp;</td>
        <td id="17_5_2">&nbsp;</td>
        <td id="17_5_3">&nbsp;</td>
        <td id="17_5_4">&nbsp;</td>
        <td id="17_5_5">&nbsp;</td>
        <td id="17_6_1">&nbsp;</td>
        <td id="17_6_2">&nbsp;</td>
        <td id="17_6_3">&nbsp;</td>
        <td id="17_6_4">&nbsp;</td>
        <td id="17_6_5">&nbsp;</td>
        <td id="17_7_1">&nbsp;</td>
        <td id="17_7_2">&nbsp;</td>
        <td id="17_7_3">&nbsp;</td>
        <td id="17_7_4">&nbsp;</td>
        <td id="17_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center" id="W18">D18</td>
        <td width="20" align="center">18</td>
        <td id="18_1_1" width="20">&nbsp;</td>
        <td id="18_1_2" width="20">&nbsp;</td>
        <td id="18_1_3" width="20">&nbsp;</td>
        <td id="18_1_4" width="20">&nbsp;</td>
        <td id="18_1_5" width="20">&nbsp;</td>
        <td id="18_2_1">&nbsp;</td>
        <td id="18_2_2">&nbsp;</td>
        <td id="18_2_3">&nbsp;</td>
        <td id="18_2_4">&nbsp;</td>
        <td id="18_2_5">&nbsp;</td>
        <td id="18_3_1">&nbsp;</td>
        <td id="18_3_2">&nbsp;</td>
        <td id="18_3_3">&nbsp;</td>
        <td id="18_3_4">&nbsp;</td>
        <td id="18_3_5">&nbsp;</td>
        <td id="18_4_1">&nbsp;</td>
        <td id="18_4_2">&nbsp;</td>
        <td id="18_4_3">&nbsp;</td>
        <td id="18_4_4">&nbsp;</td>
        <td id="18_4_5">&nbsp;</td>
        <td id="18_5_1">&nbsp;</td>
        <td id="18_5_2">&nbsp;</td>
        <td id="18_5_3">&nbsp;</td>
        <td id="18_5_4">&nbsp;</td>
        <td id="18_5_5">&nbsp;</td>
        <td id="18_6_1">&nbsp;</td>
        <td id="18_6_2">&nbsp;</td>
        <td id="18_6_3">&nbsp;</td>
        <td id="18_6_4">&nbsp;</td>
        <td id="18_6_5">&nbsp;</td>
        <td id="18_7_1">&nbsp;</td>
        <td id="18_7_2">&nbsp;</td>
        <td id="18_7_3">&nbsp;</td>
        <td id="18_7_4">&nbsp;</td>
        <td id="18_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center"  id="W19">D19</td>
        <td width="20" align="center">19</td>
        <td id="19_1_1" width="20">&nbsp;</td>
        <td id="19_1_2" width="20">&nbsp;</td>
        <td id="19_1_3" width="20">&nbsp;</td>
        <td id="19_1_4" width="20">&nbsp;</td>
        <td id="19_1_5" width="20">&nbsp;</td>
        <td id="19_2_1">&nbsp;</td>
        <td id="19_2_2">&nbsp;</td>
        <td id="19_2_3">&nbsp;</td>
        <td id="19_2_4">&nbsp;</td>
        <td id="19_2_5">&nbsp;</td>
        <td id="19_3_1">&nbsp;</td>
        <td id="19_3_2">&nbsp;</td>
        <td id="19_3_3">&nbsp;</td>
        <td id="19_3_4">&nbsp;</td>
        <td id="19_3_5">&nbsp;</td>
        <td id="19_4_1">&nbsp;</td>
        <td id="19_4_2">&nbsp;</td>
        <td id="19_4_3">&nbsp;</td>
        <td id="19_4_4">&nbsp;</td>
        <td id="19_4_5">&nbsp;</td>
        <td id="19_5_1">&nbsp;</td>
        <td id="19_5_2">&nbsp;</td>
        <td id="19_5_3">&nbsp;</td>
        <td id="19_5_4">&nbsp;</td>
        <td id="19_5_5">&nbsp;</td>
        <td id="19_6_1">&nbsp;</td>
        <td id="19_6_2">&nbsp;</td>
        <td id="19_6_3">&nbsp;</td>
        <td id="19_6_4">&nbsp;</td>
        <td id="19_6_5">&nbsp;</td>
        <td id="19_7_1">&nbsp;</td>
        <td id="19_7_2">&nbsp;</td>
        <td id="19_7_3">&nbsp;</td>
        <td id="19_7_4">&nbsp;</td>
        <td id="19_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center" id="W20">D20</td>
        <td width="20" align="center">20</td>
        <td id="20_1_1" width="20">&nbsp;</td>
        <td id="20_1_2" width="20">&nbsp;</td>
        <td id="20_1_3" width="20">&nbsp;</td>
        <td id="20_1_4" width="20">&nbsp;</td>
        <td id="20_1_5" width="20">&nbsp;</td>
        <td id="20_2_1">&nbsp;</td>
        <td id="20_2_2">&nbsp;</td>
        <td id="20_2_3">&nbsp;</td>
        <td id="20_2_4">&nbsp;</td>
        <td id="20_2_5">&nbsp;</td>
        <td id="20_3_1">&nbsp;</td>
        <td id="20_3_2">&nbsp;</td>
        <td id="20_3_3">&nbsp;</td>
        <td id="20_3_4">&nbsp;</td>
        <td id="20_3_5">&nbsp;</td>
        <td id="20_4_1">&nbsp;</td>
        <td id="20_4_2">&nbsp;</td>
        <td id="20_4_3">&nbsp;</td>
        <td id="20_4_4">&nbsp;</td>
        <td id="20_4_5">&nbsp;</td>
        <td id="20_5_1">&nbsp;</td>
        <td id="20_5_2">&nbsp;</td>
        <td id="20_5_3">&nbsp;</td>
        <td id="20_5_4">&nbsp;</td>
        <td id="20_5_5">&nbsp;</td>
        <td id="20_6_1">&nbsp;</td>
        <td id="20_6_2">&nbsp;</td>
        <td id="20_6_3">&nbsp;</td>
        <td id="20_6_4">&nbsp;</td>
        <td id="20_6_5">&nbsp;</td>
        <td id="20_7_1">&nbsp;</td>
        <td id="20_7_2">&nbsp;</td>
        <td id="20_7_3">&nbsp;</td>
        <td id="20_7_4">&nbsp;</td>
        <td id="20_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center" id="W21">D21</td>
        <td width="20" align="center">21</td>
        <td id="21_1_1" width="20">&nbsp;</td>
        <td id="21_1_2" width="20">&nbsp;</td>
        <td id="21_1_3" width="20">&nbsp;</td>
        <td id="21_1_4" width="20">&nbsp;</td>
        <td id="21_1_5" width="20">&nbsp;</td>
        <td id="21_2_1">&nbsp;</td>
        <td id="21_2_2">&nbsp;</td>
        <td id="21_2_3">&nbsp;</td>
        <td id="21_2_4">&nbsp;</td>
        <td id="21_2_5">&nbsp;</td>
        <td id="21_3_1">&nbsp;</td>
        <td id="21_3_2">&nbsp;</td>
        <td id="21_3_3">&nbsp;</td>
        <td id="21_3_4">&nbsp;</td>
        <td id="21_3_5">&nbsp;</td>
        <td id="21_4_1">&nbsp;</td>
        <td id="21_4_2">&nbsp;</td>
        <td id="21_4_3">&nbsp;</td>
        <td id="21_4_4">&nbsp;</td>
        <td id="21_4_5">&nbsp;</td>
        <td id="21_5_1">&nbsp;</td>
        <td id="21_5_2">&nbsp;</td>
        <td id="21_5_3">&nbsp;</td>
        <td id="21_5_4">&nbsp;</td>
        <td id="21_5_5">&nbsp;</td>
        <td id="21_6_1">&nbsp;</td>
        <td id="21_6_2">&nbsp;</td>
        <td id="21_6_3">&nbsp;</td>
        <td id="21_6_4">&nbsp;</td>
        <td id="21_6_5">&nbsp;</td>
        <td id="21_7_1">&nbsp;</td>
        <td id="21_7_2">&nbsp;</td>
        <td id="21_7_3">&nbsp;</td>
        <td id="21_7_4">&nbsp;</td>
        <td id="21_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center" id="W22">D22</td>
        <td width="20" align="center">22</td>
        <td id="22_1_1" width="20">&nbsp;</td>
        <td id="22_1_2" width="20">&nbsp;</td>
        <td id="22_1_3" width="20">&nbsp;</td>
        <td id="22_1_4" width="20">&nbsp;</td>
        <td id="22_1_5" width="20">&nbsp;</td>
        <td id="22_2_1">&nbsp;</td>
        <td id="22_2_2">&nbsp;</td>
        <td id="22_2_3">&nbsp;</td>
        <td id="22_2_4">&nbsp;</td>
        <td id="22_2_5">&nbsp;</td>
        <td id="22_3_1">&nbsp;</td>
        <td id="22_3_2">&nbsp;</td>
        <td id="22_3_3">&nbsp;</td>
        <td id="22_3_4">&nbsp;</td>
        <td id="22_3_5">&nbsp;</td>
        <td id="22_4_1">&nbsp;</td>
        <td id="22_4_2">&nbsp;</td>
        <td id="22_4_3">&nbsp;</td>
        <td id="22_4_4">&nbsp;</td>
        <td id="22_4_5">&nbsp;</td>
        <td id="22_5_1">&nbsp;</td>
        <td id="22_5_2">&nbsp;</td>
        <td id="22_5_3">&nbsp;</td>
        <td id="22_5_4">&nbsp;</td>
        <td id="22_5_5">&nbsp;</td>
        <td id="22_6_1">&nbsp;</td>
        <td id="22_6_2">&nbsp;</td>
        <td id="22_6_3">&nbsp;</td>
        <td id="22_6_4">&nbsp;</td>
        <td id="22_6_5">&nbsp;</td>
        <td id="22_7_1">&nbsp;</td>
        <td id="22_7_2">&nbsp;</td>
        <td id="22_7_3">&nbsp;</td>
        <td id="22_7_4">&nbsp;</td>
        <td id="22_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center" id="W23">D23</td>
        <td width="20" align="center">23</td>
        <td id="23_1_1" width="20">&nbsp;</td>
        <td id="23_1_2" width="20">&nbsp;</td>
        <td id="23_1_3" width="20">&nbsp;</td>
        <td id="23_1_4" width="20">&nbsp;</td>
        <td id="23_1_5" width="20">&nbsp;</td>
        <td id="23_2_1">&nbsp;</td>
        <td id="23_2_2">&nbsp;</td>
        <td id="23_2_3">&nbsp;</td>
        <td id="23_2_4">&nbsp;</td>
        <td id="23_2_5">&nbsp;</td>
        <td id="23_3_1">&nbsp;</td>
        <td id="23_3_2">&nbsp;</td>
        <td id="23_3_3">&nbsp;</td>
        <td id="23_3_4">&nbsp;</td>
        <td id="23_3_5">&nbsp;</td>
        <td id="23_4_1">&nbsp;</td>
        <td id="23_4_2">&nbsp;</td>
        <td id="23_4_3">&nbsp;</td>
        <td id="23_4_4">&nbsp;</td>
        <td id="23_4_5">&nbsp;</td>
        <td id="23_5_1">&nbsp;</td>
        <td id="23_5_2">&nbsp;</td>
        <td id="23_5_3">&nbsp;</td>
        <td id="23_5_4">&nbsp;</td>
        <td id="23_5_5">&nbsp;</td>
        <td id="23_6_1">&nbsp;</td>
        <td id="23_6_2">&nbsp;</td>
        <td id="23_6_3">&nbsp;</td>
        <td id="23_6_4">&nbsp;</td>
        <td id="23_6_5">&nbsp;</td>
        <td id="23_7_1">&nbsp;</td>
        <td id="23_7_2">&nbsp;</td>
        <td id="23_7_3">&nbsp;</td>
        <td id="23_7_4">&nbsp;</td>
        <td id="23_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center" id="W24">D24</td>
        <td width="20" align="center">24</td>
        <td id="24_1_1" width="20">&nbsp;</td>
        <td id="24_1_2" width="20">&nbsp;</td>
        <td id="24_1_3" width="20">&nbsp;</td>
        <td id="24_1_4" width="20">&nbsp;</td>
        <td id="24_1_5" width="20">&nbsp;</td>
        <td id="24_2_1">&nbsp;</td>
        <td id="24_2_2">&nbsp;</td>
        <td id="24_2_3">&nbsp;</td>
        <td id="24_2_4">&nbsp;</td>
        <td id="24_2_5">&nbsp;</td>
        <td id="24_3_1">&nbsp;</td>
        <td id="24_3_2">&nbsp;</td>
        <td id="24_3_3">&nbsp;</td>
        <td id="24_3_4">&nbsp;</td>
        <td id="24_3_5">&nbsp;</td>
        <td id="24_4_1">&nbsp;</td>
        <td id="24_4_2">&nbsp;</td>
        <td id="24_4_3">&nbsp;</td>
        <td id="24_4_4">&nbsp;</td>
        <td id="24_4_5">&nbsp;</td>
        <td id="24_5_1">&nbsp;</td>
        <td id="24_5_2">&nbsp;</td>
        <td id="24_5_3">&nbsp;</td>
        <td id="24_5_4">&nbsp;</td>
        <td id="24_5_5">&nbsp;</td>
        <td id="24_6_1">&nbsp;</td>
        <td id="24_6_2">&nbsp;</td>
        <td id="24_6_3">&nbsp;</td>
        <td id="24_6_4">&nbsp;</td>
        <td id="24_6_5">&nbsp;</td>
        <td id="24_7_1">&nbsp;</td>
        <td id="24_7_2">&nbsp;</td>
        <td id="24_7_3">&nbsp;</td>
        <td id="24_7_4">&nbsp;</td>
        <td id="24_7_5">&nbsp;</td>
      </tr>
      <tr>
        <td width="220" align="center" id="W25">D25</td>
        <td width="20" align="center">25</td>
        <td id="25_1_1" width="20">&nbsp;</td>
        <td id="25_1_2" width="20">&nbsp;</td>
        <td id="25_1_3" width="20">&nbsp;</td>
        <td id="25_1_4" width="20">&nbsp;</td>
        <td id="25_1_5" width="20">&nbsp;</td>
        <td id="25_2_1">&nbsp;</td>
        <td id="25_2_2">&nbsp;</td>
        <td id="25_2_3">&nbsp;</td>
        <td id="25_2_4">&nbsp;</td>
        <td id="25_2_5">&nbsp;</td>
        <td id="25_3_1">&nbsp;</td>
        <td id="25_3_2">&nbsp;</td>
        <td id="25_3_3">&nbsp;</td>
        <td id="25_3_4">&nbsp;</td>
        <td id="25_3_5">&nbsp;</td>
        <td id="25_4_1">&nbsp;</td>
        <td id="25_4_2">&nbsp;</td>
        <td id="25_4_3">&nbsp;</td>
        <td id="25_4_4">&nbsp;</td>
        <td id="25_4_5">&nbsp;</td>
        <td id="25_5_1">&nbsp;</td>
        <td id="25_5_2">&nbsp;</td>
        <td id="25_5_3">&nbsp;</td>
        <td id="25_5_4">&nbsp;</td>
        <td id="25_5_5">&nbsp;</td>
        <td id="25_6_1">&nbsp;</td>
        <td id="25_6_2">&nbsp;</td>
        <td id="25_6_3">&nbsp;</td>
        <td id="25_6_4">&nbsp;</td>
        <td id="25_6_5">&nbsp;</td>
        <td id="25_7_1">&nbsp;</td>
        <td id="25_7_2">&nbsp;</td>
        <td id="25_7_3">&nbsp;</td>
        <td id="25_7_4">&nbsp;</td>
        <td id="25_7_5">&nbsp;</td>
      </tr>
    </table>
    <p>  12节：0810-0950&nbsp;&nbsp;&nbsp;&nbsp;34节：1000-1140&nbsp;&nbsp;&nbsp;&nbsp;56节：1430-1610&nbsp;&nbsp;&nbsp;&nbsp;
    78节：1620-1800&nbsp;&nbsp;&nbsp;&nbsp;晚：1900-2040</p>
    
    </body>
    </html>
    '''
    # print(str(days), str(day_jieci))
    table_template = table_template.replace('#back__color', backcolor)
    table_template = table_template.replace('#td_list', str(td_list))
    webview.load_html(table_template)

def loadthread(backcolor='', td_list=''):
    t = threading.Thread(target=loadhtml, args=(backcolor, td_list))
    t.setDaemon(True)
    t.start()
    t.join(2)

def dateRange(beginDate, endDate):
    # 用于获得起始日期之间的日期列表
    dates = []
    dt = datetime.datetime.strptime(beginDate, "%Y-%m-%d")
    date = beginDate[:]
    while date <= endDate:
        dates.append(date)
        dt = dt + datetime.timedelta(1)
        date = dt.strftime("%Y-%m-%d")
    return dates

def get_date_period_of_weeks(start_date_is_monday, lastdate_is_sunday):
    # 用于获得起始日期之间，每周次的日期区间列表
    weekday_start = datetime.datetime.strptime(start_date_is_monday, "%Y-%m-%d").weekday()
    weekday_last = datetime.datetime.strptime(lastdate_is_sunday, "%Y-%m-%d").weekday()
    if weekday_start != 0 or weekday_last != 6:
        # 没有输入正确的参数，起始结束日期必须是周一和周日
        raise ValueError
    else:
        periods = []
        countday = len(dateRange(start_date_is_monday, lastdate_is_sunday))
        # countday必然是7的倍数，是日期区间包含的天数
        countweek = countday / 7
        # countweek>=1，是日期区间包含的周数
        if countweek > 25:
            # 输入了错误的时间区间，标准课表每学期最多不能超过25周
            raise ValueError
        countw = 0
        dt = datetime.datetime.strptime(start_date_is_monday, "%Y-%m-%d")
        while countw < countweek:
            countw = countw + 1
            dt_other = dt + datetime.timedelta(countw * 7 - 7)
            dayfront = str(dt_other) + '至'
            dayend = str(dt_other + datetime.timedelta(6))
            periods.append(dayfront.replace(' 00:00:00', '') + dayend.replace(' 00:00:00', ''))
        return periods

class Api:
    def __init__(self, start_date_mustbe_monday, lastdate_mustbe_sunday):
        self._start_date_is_monday = start_date_mustbe_monday
        self._lastdate_is_sunday = lastdate_mustbe_sunday

    def getWeeks(self, params=''):
        response = {'message': get_date_period_of_weeks(self._start_date_is_monday, self._lastdate_is_sunday)}
        # print(response)
        return response

