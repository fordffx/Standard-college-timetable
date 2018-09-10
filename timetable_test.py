import threading
import webview
import timetable
import datetime

if __name__ == '__main__':
    Paccessdb = "kb.mdb"

    psql_room_use_plan = "select 日期,节次 from kb where 上课地点 = 'z400' and cdate(日期)>#" + "2018-02-19" + "# and cdate(日期)<#" + "2018-07-29" + "#"
    room_use_plan = timetable.read_kebiao(Paccessdb, '', psql_room_use_plan)
    #print(room_use_plan)
    #根据查询语句获取教室使用日期和当天节次的列表
    #下面将2个列表整理格式
    days = []
    day_jieci=[]
    someday=0
    daynum=len(room_use_plan)
    while someday < daynum:
        s=room_use_plan[someday][0]
        s=str(s).replace(' 00:00:00','')
        days.append(s)
        day_jieci.append(room_use_plan[someday][1])
        someday = someday + 1
    #print(days,day_jieci)
    #获取学期中周次对应的区间范围列表
    api = timetable.Api('2018-02-19', '2018-07-29')
    message_dict = api.getWeeks()
    week_list = message_dict['message']
    #生成图示中对应的表格td清单，传递给html中的javascript绘画底色使用
    L_i=0
    td_list=[]
    while L_i < len(days):
        L_j = 0
        while L_j < len(week_list):
            _leftdate = week_list[L_j].split('至')[0]
            _rightdate = week_list[L_j].split('至')[1]
            if days[L_i] >= _leftdate and days[L_i] <= _rightdate:
                _weekday = datetime.datetime.strptime(days[L_i], "%Y-%m-%d").weekday()
                def _return(x):
                    return x
                _jc={'12':_return('1'),'34':_return('2'),'56':_return('3'),'78':_return('4'),'晚':_return('5')}
                _tempvar = str(L_j+1) + '_' + str(_weekday+1) + '_' + _jc[day_jieci[L_i]]
                td_list.append(_tempvar)
                break
            L_j = L_j + 1
        L_i=L_i+1
    #print(td_list)

    t = threading.Thread(target=timetable.loadthread, args=('seashell',td_list))
    # mediumseagreen
    t.start()
    webview.create_window('xxx教室本学期使用计划', js_api=api, width=1040, height=615, debug=True)
