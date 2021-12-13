import  pandas as pd
import tkinter as tk
from tkinter import filedialog
import tkinter.messagebox

def copy_excel_row():
    # Use a breakpoint in the code line below to debug your script.
    root = tk.Tk()
    root.withdraw()
    filepath = filedialog.askopenfilename()
    print(filepath)
    outfile = f"{filepath.replace('.xlsx','').replace('.xls','')}拆分后文件.xlsx"
    old_data = pd.read_excel(filepath)
    try:
        old_data['数量'] = old_data['数量'].fillna(1)
    except Exception as e:
        print('error:',e)
        if '数量' in str(e):
            old_data['数量']=0
        else:
            tkinter.messagebox.showerror('未知错误', f'错误代码：{e}')
    # print(old_data)
    old_data_list = old_data.to_dict('records')
    # print(old_data_dict)
    new_data = []
    # c = 1
    for i in old_data_list:
        # c+= 1
        # print(i.get('数量',0))
        # print(type(i.get('数量',0)))
        # count = 0 if i.get('数量',0) == np.NaN else int(i.get('数量',0))
        # print(count)
        for j in range(int(i.get('数量',1))):
            new_data.append(i)
        # if c >5:
        #     break
    res_data = pd.DataFrame(new_data)
    res_data = res_data.drop(columns='数量')
    res_data.to_excel(outfile)
    print('处理完毕')
    tkinter.messagebox.showinfo('成功', '处理完毕，处理后文件与原文件保存位置相同！')
    # print(old_data)
    # print(old_data.loc[1])
    # new_data = old_data.drop(index=old_data.index)
    # print(new_data)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    copy_excel_row()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
