import pandas as pd
import os
import tkinter as tk
# from tkinter import ttk
from tkinter import filedialog as fd

# 店小蜜需要的数据
dianxiaomi_col = ['*订单号', '*店铺账号', '*sku',	"属性(可填写SKU尺寸、颜色等)", '*数量（大于0的整数）',
                  '*单价', '总运费', '币种（默认USD）', '*买家姓名',	'*地址1',	'地址2',	'*城市',	'*省/州',	'*国家二字码',	'*邮编',	'电话',]
# 获取onbuy中需要的数据
onbuy_col = ['Order Number', 'SKU', 'Quantity', 'Product Unit Price', 'Customer', 'Delivery Address Name',
             'Delivery Address Line 1', 'Delivery Address Line 2', 'Delivery Address Line 3', 'Delivery Address Town', 'Delivery Address County', 'Site', 'Delivery Address Postcode']

# 处理customer字段藏手机号
def split_customer(str):
    return str.split(',')[1]

# 获取excel数据
def get_onbuy_data(path=''):
    if not path:
        return
    # 店面名
    s_path = path.split('/')
    f_dm = s_path[-1] or '未设置默认店名'
    # print('====', path, s_path[-1])

    if path.endswith('.xls'):
        # 读取excel
        df = pd.read_excel(rf'{path}', usecols=onbuy_col,
                           converters={'Customer': split_customer})
    elif path.endswith('.csv'):
        df = pd.read_csv(rf'{path}', usecols=onbuy_col,
                         converters={'Customer': split_customer}, encoding='utf-8')
    # df = pd.concat([df1, df2])
    # print(df, '\n  -------')
    # df_li = df.values.tolist()
    # print(df_li, '\n -----------')

    # 获取数据长度填充空白
    df_length = len(df)
    # 空表头
    empt = ''*df_length
    # 处理address、Delivery Address County
    town = list(df['Delivery Address Town'])
    county = list(df['Delivery Address County'])
    address2 = list(df['Delivery Address Line 2'])
    address3 = list(df['Delivery Address Line 3'])
    # print(df.loc[1, []], '-=======----')
    # print(town)

    for i in range(df_length):
        # town是否多个
        town_arr = town[i].split(',')
        if len(town_arr) >1:
            # town：城市有','没后续名字
            if town_arr[-1].isspace():
                town[i] = town_arr[0] # 去除','
                town_arr[-1] = town_arr[0]
            else:
                # 地址3是否空
                if pd.isnull(address3[i]):
                    address3[i] = town_arr[0]
                    town[i] = town_arr[-1]
                
        # 省/州是否空
        if pd.isnull(county[i]):
            county[i] = town_arr[-1]

        # 地址3和地址2处理叠加不为空
        if pd.isnull(address2[i]):
            address2[i] = ''
        if pd.isnull(address3[i]):
            address3[i] = ''
        address2[i] = (address2[i]+','+address3[i]).strip(',')

    # 处理表头数据
    dianxiaomi_li = {
        '*订单号': list(df['Order Number']),
        '*店铺账号': ['手工订单']*df_length,
        '*sku': list(df['SKU']),
        "属性(可填写SKU尺寸、颜色等)": [f_dm]*df_length,
        '*数量（大于0的整数）': list(df['Quantity']),
        '*单价': list(df['Product Unit Price']),
        '总运费': empt,
        '币种（默认USD）': empt,
        '*买家姓名': list(df['Delivery Address Name']),
        '*地址1': list(df['Delivery Address Line 1']),
        '地址2': address2,
        '*城市': town,
        '*省/州': county,
        '*国家二字码': ['UK']*df_length,
        '*邮编': list(df['Delivery Address Postcode']),
        '电话': list(df['Customer']),
        '手机': empt,
        'E-mail': empt,
        '买家税号': empt,
        '门牌号': empt,
        '公司名': empt,
        '订单备注': empt,
        '图片网址': empt,
        '售出链接': empt,
        '中文报关名': empt,
        '英文报关名': empt,
        '申报金额（USD）': empt,
        '申报重量（g）': empt,
        '材质': empt,
        '用途': empt,
        '海关编码': empt,
        '报关属性': empt,
        '卖家税号（IOSS）': empt,
    }
    # print(dianxiaomi_li, '\n ---------')
    # print(pd.DataFrame(dianxiaomi_li))
    return pd.DataFrame(dianxiaomi_li)
    # df = pd.DataFrame(df_li, columns=dianxiaomi_col)


# 生成目标文件
def get_dist_data(file_list = []):
    count_data = pd.DataFrame()
    # file_list = get_file()
    if not len(file_list):
        print('没有需要转换的csv/xls文件')
        return
    print(file_list)
    o_name = file_list[0].split('/')[-1][:-4]
    print('o_name',o_name)
    for i in file_list:
        count_data = pd.concat([count_data, get_onbuy_data(i)])
        # 重命名文件
        fname = os.path.basename(i)
        # os.rename(i, i.replace(fname, '(源数据)'+fname))

    # print('====', count_data)
    count_data.to_excel(o_name+"店小秘订单.xls", index=False)
    print(f'生成{o_name}店小秘订单文件')


# 扫描xls文件
def get_file(path='./'):
    print('获取目录.....')
    files = os.listdir(path)
    # files = os.walk(path)
    xls = []
    if not len(files):
        print('目录为空，不执行操作\n')

    for f in files:
        # f_name = f.split('.')
        # if (f_name[-1] == 'csv' or f_name[-1] == 'xls'):
        if f.endswith('csv') and ('源数据' not in f):
            xls.append(f)
            print('待处理xls文件：', os.path.join(path, f))
    return xls


# 判断文件是否以打开
def file_is_openState(file_path):
    try:
        print(open(file_path, "w"))
        return False
    except Exception as e:
        if ("[Errno 13] Permission denied" in str(e)):
            print("文件已打开!")
            return True
        else:
            return False
        
# 多选择文件
def select_file():
    filepath = fd.askopenfilenames(filetypes=[('.CSV','.csv'), ('.XLS','.xls')]) # 选择打开什么文件，返回文件名
    if not len(filepath): 
        return
    
    file_list.delete(1.0,'end')
    fstr = '读取的文件:\n '
    xls = []
    for f in filepath:
        fstr = fstr + '\n' + f
        if f.endswith('csv') and ('源数据' not in f):
            xls.append(f)
            print('待处理csv/xls文件：', f)

    # 获取目标数据
    get_dist_data(xls)

    # 'insert'表示光标插入位置
    file_list.insert('insert',fstr)

# 清除显示
def clear_console():
    file_list.delete(1.0,'end')

# 删除列表中文件
def delete_file():
    context = file_list.get(1.0,'end').split('\n')
    for c in context:
        if os.path.isfile(c):
            os.remove(c)
    # fname = os.path.basename(i)
    # os.rename(i, i.replace(fname, '(源数据)'+fname))
    file_list.delete(1.0,'end')
    file_list.insert('insert','删除了！！！')

if __name__ == '__main__':
    # get_dist_data()
    # input('按回车退出…')
    root  = tk.Tk()  # 创建窗口对象
    root.title('onbuy2店小蜜')
    root.geometry('800x600')
    # root.resizable(False, False) # 规定窗口不可缩放

    f_lable = tk.Label(root, text='onbuy订单转店小蜜', font=('bold',20), justify='center',padx=20,pady=20).pack()
    fm1 = tk.Frame(root)
    select_btn = tk.Button(fm1, text ="选择文件", font='20', command = select_file)
    clear_btn = tk.Button(fm1, text ="清除显示", font='20', command = clear_console)
    delete_btn = tk.Button(fm1, text ="删除列表中的文件", font='20', command = delete_file)
    file_list = tk.Text(root, width=80, height=20 ,padx=0,font=20)
    fm1.pack()
    select_btn.pack(side= 'left',padx=10,pady=10)
    clear_btn.pack(side= 'left',padx=10,pady=10)
    delete_btn.pack(side= 'left',padx=10,pady=10)
    file_list.pack()

    root.mainloop()  # 进入消息循环

