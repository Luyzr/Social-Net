import pandas as pd
import openpyxl as op

def showsome():
    # 小表白
    print('来自爱你的老公酱(｡･ω･｡)ﾉ♡')

def gettime(timelist):
    # 根据年月日时分得到一个整数类型的数字表示时间
    return 24*60*(365*timelist[0] + 31*timelist[1] + timelist[2]) + 60*timelist[3] + timelist[4]

def divide(data, limits):
# 根据最早的转发时间确定该部分数据的阶段归属并对其进行分类
    divided = {1:[], 2:[], 3:[], 4:[]}
    formal = ''
    limit_t = list(map(gettime, limits))
    pubt = 0
    temp = [[]]
    for row in data.itertuples():
        owner = row[1]
        affected = row[2]
        if type(row[4]) == str:
            t = list(map(int, row[4].split()[0].split('-')))
            _time = list(map(int, row[4].split()[1].split(':')))
            t.extend(_time)
            t = gettime(t)
            notes = row[6]
            if formal != owner and formal != '':
                if pubt <= limit_t[0]:
                    divided[1].append(temp)
                elif pubt <= limit_t[1]:
                    divided[2].append(temp)
                elif pubt <= limit_t[2]:
                    divided[3].append(temp)
                else:
                    divided[4].append(temp)
                temp = [[owner, affected, notes]]
                formal = owner
                pubt = t
            else:
                formal = owner
                if temp == [[]]:
                    temp = [[owner, affected, notes]]
                else:
                    temp.append([owner, affected, notes])
                if pubt > t:
                    pubt = t
    return divided
 
def countnum(isshow, groups):
    # 获得单阶段所有博文的平均转发量
    maxnum = 0
    minnum = 1e7
    averagenum = 0
    count = 0
    numb = []
    for group_idx in range(len(groups)):
        num = len(groups[group_idx])
        numb.append(num)
        count += num
        # print('{} has {} transmits'.format(data[groups][0][0], num))
    averagenum = sum(numb)/len(groups)
    maxnum = max(numb)
    minnum = min(numb)
    middlenum = sorted(numb)[len(numb)//2]
    if isshow:
        print('In this period, there are {} blogs, the transmit number is: {},  the maxnum is {}, the minnum is {}, the averagenum is {}, the middlenum is {}\n'.format(len(groups), sorted(numb), maxnum, minnum, averagenum, middlenum))
    return averagenum

def selectgroup(divided):
    # 将转发量小于该阶段所有博文的平均转发量的组删除
    for i in range(4):
        period = i + 1
        groups = divided[period]
        averagenum = countnum(True, groups)
        newgroups = []
        for group_idx in range(len(groups)):
            group = groups[group_idx]
            if len(group) > averagenum:
                newgroups.append(group)
        # newavg = countnum(False, newgroups)
        divided[period] = newgroups
    return divided
                    
def selectrow(group):
    # 限定每组最多只能容纳的转发量：如果转发用户被别人转发了那一定保存，其它的填充数量至原转发量的1/10
    lim = len(group)//10
    newgroup = []
    count = 0
    for row in group:
        if type(row[2]) == str and row[2].startswith('@'):
            newgroup.append(row)
            count += 1
    if count < lim:
        for row in group:
            if count < lim:
                if type(row[2]) != str:
                    newgroup.append(row)
                    count += 1
    return newgroup
        
def selectdata(divided):
    # 数据删减
    divided = selectgroup(divided)
    for i in range(4):
        period = i + 1
        groups = divided[period]
        for group_idx in range(len(groups)):
            newgroup = selectrow(groups[group_idx])
            groups[group_idx] = newgroup
    return divided

def list2txt(nou, ans, k, txtpath):
    # 将结果存储为可供Ucinet使用的数据格式
    with open('result/' + txtpath, 'w', encoding='utf-8') as f:
            f.write('dl n={}\nlabels:\n'.format(nou-1))
            for i in range(len(ans)-1):
                f.write('{},'.format(ans[i]))
            f.write('{}\ndata:\n'.format(ans[-1]))
            for i in range(1, len(k)):
                for j in range(1, len(k[0])):
                    f.write('{} '.format(k[i][j]))
                f.write('\n')

def newlist2txt(nou, ans, new_ans, k, txtpath):
    # 由于我的Ucinet没办法显示中文，故将原用户名换为英文标记，再配套一个检索文件
    with open('result/' + txtpath, 'w', encoding='utf-8') as f:
            f.write('dl n={}\nlabels:\n'.format(nou-1))
            for i in range(len(new_ans)-1):
                f.write('{},'.format(new_ans[i]))
            f.write('{}\ndata:\n'.format(new_ans[-1]))
            for i in range(1, len(k)):
                for j in range(1, len(k[0])):
                    f.write('{} '.format(k[i][j]))
                f.write('\n')
    with open('result/parallelism_' + txtpath, 'w', encoding='utf-8') as f:
            f.write('nickname            realname\n')
            for i in range(len(new_ans)-1):
                f.write('{:<20}{:<20}\n'.format(new_ans[i], ans[i]))

def list2xls(k, xlspath):
    # 将列表转为xlsx文件
    wb = op.Workbook()
    ws = wb.active
    #将数据写入第 i 行，第 j 列
    i = 0
    for i in range(len(k)):
        ws.append(k[i])
        
    wb.save('result/' + xlspath) #保存文件

def getgephi(period, k):
    # 将数据整理成可供gephi使用的数据格式
    weight = [5, 20, 5, 3]
    gephi = [['Source', 'Target', 'Weight']]
    n = len(k)
    for i in range(1, n):
        for j in range(1, n):
            if k[i][j] >= weight[period]:
                gephi.append([k[i][0], k[0][j], k[i][j]])
    return gephi

def getdata(periods ,divided, startperiod):
    # 获得每个阶段的社会网络矩阵
    for period in range(periods):
        dic = {}
        ans = []
        new_ans = [] # 用于标定博主和用户
        ownernum = 0
        affectednum = 0
        txtpath = 'data_{}.txt'.format(period)
        xlspath = 'data_{}.xlsx'.format(period)
        for group in divided[period + startperiod]:
            for row in group:
                owner = row[0]
                affected = row[1]
                notes = row[2]
                if dic.get(owner) == None:
                    dic[owner] = 1
                    ans.append(owner)

                    ownernum += 1
                    new_ans.append('owner{}'.format(ownernum))

                if dic.get(affected) == None:
                    dic[affected] = 1
                    ans.append(affected)

                    affectednum += 1
                    new_ans.append('affected{}'.format(affectednum))

                if type(notes) == str:
                    notel = notes.split(',')
                    for note in notel:
                        if note.startswith('@'):
                            n = note.split('@')[1]
                            if n != affected and n != owner:
                                if dic.get(n) == None:
                                    dic[n] = 1
                                    ans.append(n)

                                    affectednum += 1
                                    new_ans.append('affected{}'.format(affectednum))

        nou = len(ans) + 1
        k = [[0]*nou for i in range(nou)]
        print('Number of accounts in period {}: {}\n Processing...'.format(period+1, len(ans)))
        for j in range(nou-1):
            k[0][j+1] = ans[j]
            k[j+1][0] = ans[j]

        for group in divided[period + startperiod]:
            for row in group:
                owner = ans.index(row[0]) + 1
                affected = ans.index(row[1]) + 1
                k[owner][affected] += 1
                notes = row[2]
                if type(notes) == str:
                    notel = notes.split(',')
                    for note in notel:
                        if note.startswith('@'):
                            n = note.split('@')[1]
                            if n != affected and n != owner:
                                nn = ans.index(n) + 1
                                k[nn][affected] += 1
                                k[owner][nn] += 1
        # list2txt(nou, ans, k, txtpath)
        gephi = getgephi(period, k)
        list2xls(gephi, xlspath)
        # newlist2txt(nou, ans, new_ans, k, txtpath)
                

if __name__ == '__main__':
    showsome()
    path = 'test.xlsx'
    # 待处理的Excel文件
    limits = [[18, 10, 28, 17, 50], [18, 11, 1, 0, 0], [18, 11, 3, 0, 0]]
    # 5个时间节点[年, 月, 日, 时, 分]
    data = pd.read_excel(path,sheet_name = 0, names=['owner', 'affected', 'famous', 'date', 'link', 'notes'])
    print('Path of Excel:{}\nTime for dividing:{}\nDeviding...'.format(path, limits))
    divided = divide(data, limits)
    print('Before select:')
    divided = selectdata(divided)
    print('After select:')
    for i in range(4):
        period = i + 1
        avg = countnum(True, divided[period])
    # countnum(divided)
    getdata(4, divided, 1)
    # getdata(跨越阶段数，数据，开始阶段)