from tkinter import *
from docx import Document
import os

# def print_hi(name):
#     # 在下面的代码行中使用断点来调试脚本。
#     print(f'Hi, {name}')  # 按 Ctrl+F8 切换断点。
#
#
# # 按间距中的绿色按钮以运行脚本。
# if __name__ == '__main__':
#     print_hi('PyCharm')

# https://www.codegrepper.com/code-examples/python/python+tkinter+text+get

def main():

    def focus_next_window(event):
        event.widget.tk_focusNext().focus()
        return "break"

    def shuttle_text(shuttle):
        t = ''
        for i in shuttle:
            t += i.text
        return t

    def docx_replace(doc, data):
        for key in data:
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if key in cell.text:
                            cell.text = cell.text.replace(key, data[key])

            for p in doc.paragraphs:

                begin = 0
                for end in range(len(p.runs)):

                    shuttle = p.runs[begin:end + 1]

                    full_text = shuttle_text(shuttle)
                    if key in full_text:
                        # print('Replace：', key, '->', data[key])
                        # print([i.text for i in shuttle])

                        # find the begin
                        index = full_text.index(key)
                        # print('full_text length', len(full_text), 'index:', index)
                        while index >= len(p.runs[begin].text):
                            index -= len(p.runs[begin].text)
                            begin += 1

                        shuttle = p.runs[begin:end + 1]

                        # do replace
                        # print('before replace', [i.text for i in shuttle])
                        if key in shuttle[0].text:
                            shuttle[0].text = shuttle[0].text.replace(key, data[key])
                        else:
                            replace_begin_index = shuttle_text(shuttle).index(key)
                            replace_end_index = replace_begin_index + len(key)
                            replace_end_index_in_last_run = replace_end_index - len(shuttle_text(shuttle[:-1]))
                            shuttle[0].text = shuttle[0].text[:replace_begin_index] + data[key]

                            # clear middle runs
                            for i in shuttle[1:-1]:
                                i.text = ''

                            # keep last run
                            shuttle[-1].text = shuttle[-1].text[replace_end_index_in_last_run:]

                        # print('after replace', [i.text for i in shuttle])

                        # set begin to next
                        begin = end

    def word_gen():
        new_name = textbox_name.get("1.0", "end-1c")
        user_id = textbox_uid.get("1.0", "end-1c")
        new_date = textbox_date.get("1.0", "end-1c")
        new_addr = textbox_addr.get("1.0", "end-1c")
        new_currency = textbox_curr.get("1.0", "end-1c")
        new_city = textbox_city.get("1.0", "end-1c")
        new_laws = textbox_laws.get("1.0", "end-1c")
        new_effect = textbox_effe.get("1.0", "end-1c")
        new_notice = textbox_noti.get("1.0", "end-1c")
        new_compen = textbox_comp.get("1.0", "end-1c")

        doc_name = user_id
        password = user_id

        print(new_name, new_date, new_addr, new_city, new_laws, new_currency, new_effect, new_notice,
              new_compen + '获取正确!!')

        docx_replace(doc, dict({old_date: new_date, old_name: new_name, old_addr: new_addr, old_city: new_city,
                                old_laws: new_laws, old_currency: new_currency, old_effect: new_effect,
                                old_notice: new_notice, old_compen: new_compen}))


        #  todo 客户单号作为后缀
        #  filename = 'Independent_Contractor_Agreement-' + 客户单号 + '.docx'
        filename = 'Independent_Contractor_Agreement_UserID' + doc_name + '.docx'
        filepath = './'
        doc.save(filepath + filename)

        #  加密word，密码是user ID

        #  打开word
        os.system('start' + filepath)

    win = Tk()
    win.title('WeProtect内部系统 Inner Contract Gen System')
    win.geometry("800x300")

    #  grid label
    label_name = ['User(用户):', 'UID(用户ID)', 'Date(日期):', 'Address(地址):', 'Currency(货币):', 'City,Zip-code(城市,邮编):',
                  'Laws of the State(管辖权):', 'Effectiveness(时效(月份)):', 'Days of Notice Ahead(提前几日通知):',
                  'Compensation(酬劳条款):']

    i = 0
    for c in label_name:
        Label(text=c).grid(column=0, row=i)
        i = i + 1

    textbox_name = Text(win, height=1)
    textbox_name.grid(column=1, row=0)
    textbox_name.bind("<Tab>", focus_next_window)
    textbox_uid = Text(win, height=1)
    textbox_uid.grid(column=1, row=1)
    textbox_uid.bind("<Tab>", focus_next_window)
    textbox_date = Text(win, height=1)
    textbox_date.grid(column=1, row=2)
    textbox_date.bind("<Tab>", focus_next_window)
    textbox_addr = Text(win, height=1)
    textbox_addr.grid(column=1, row=3)
    textbox_addr.bind("<Tab>", focus_next_window)
    textbox_curr = Text(win, height=1)
    textbox_curr.grid(column=1, row=4)
    textbox_curr.bind("<Tab>", focus_next_window)
    textbox_city = Text(win, height=1)
    textbox_city.grid(column=1, row=5)
    textbox_city.bind("<Tab>", focus_next_window)
    textbox_laws = Text(win, height=1)
    textbox_laws.grid(column=1, row=6)
    textbox_laws.bind("<Tab>", focus_next_window)
    textbox_effe = Text(win, height=1)
    textbox_effe.grid(column=1, row=7)
    textbox_effe.bind("<Tab>", focus_next_window)
    textbox_noti = Text(win, height=1)
    textbox_noti.grid(column=1, row=8)
    textbox_noti.bind("<Tab>", focus_next_window)
    textbox_comp = Text(win, height=3)
    textbox_comp.grid(column=1, row=9)
    textbox_comp.bind("<Tab>", focus_next_window)

    old_date = 'AIdate'
    old_name = 'AIname'
    old_addr = 'AIroadaddress'
    old_city = 'AIcityandcode'
    old_laws = 'AIlawsstate'
    old_currency = 'AIcurrency'
    old_effect = 'AIeffect'
    old_notice = 'AInotice'
    old_compen = 'AIcompen'

    doc = Document('./basefile/Contract.docx')

    #  word生成 按钮
    # Button(win, text='合同生成', command=word_gen).place(x=90, y=380)
    Button(win, text='合同生成(Gen Contract)', command=word_gen).grid(column=0, row=14)
    win.iconbitmap('./logo.ico')
    win.mainloop()  # 进入消息循环


if __name__ == '__main__':
    main()
