import tkinter
from tkinter import ttk
from tkinter.ttk import Treeview
import re
import os
from tkinter import messagebox
import subprocess


DEFAULT_VALUE = ' '


class Gui:
    def __init__(self):
        self.root = tkinter.Tk()
        self.root.title('allpairs - python GUI')
        '''修复官方库设置背景色无效的bug'''
        style = ttk.Style(self.root)
        def fixed_map(option):
            # Returns the style map for 'option' with any styles starting with
            # ("!disabled", "!selected", ...) filtered out
            # style.map() returns an empty list for missing options, so this should
            # be future-safe
            return [elm for elm in style.map("Treeview", query_opt=option) if elm[:2] != ("!disabled", "!selected")]
        style.map(
            "Treeview",
            foreground=fixed_map("foreground"),
            background=fixed_map("background")
        )
        
        menu = tkinter.Menu(self.root)
        menu.add_command(label='增加列', command=self.add_column)
        menu.add_command(label='插入行', command=self.add_row)
        menu.add_command(label='编辑列名', command=self.edit_column)
        menu.add_command(label='生成正交用例', command=self.make_testcase)
        menu.add_command(label='生成全排列', command=self.make_all_permutations)
        self.column = ['1','2','3']
        self.tree = Treeview(self.root, column=self.column, show="headings")
        for column in self.column:
            self.tree.heading(column, text=column)  # 设置表头
        self.tree.grid()
        self.tree.tag_configure('self', background='#d9ead3')
        self.tree.insert('', 'end', values=[DEFAULT_VALUE for i in range(len(self.column))], tags=('self',))
        self.root.config(menu=menu)
        self.index = 4
        self.tree.bind('<Button-1>', lambda event: self.edit_column_content(event))
        self.tree.bind('<Button-3>', lambda event: self.pop_menu(event))
        # self.tree.bind('<Button-3>')
        
    def main(self):
        self.root.mainloop()
        
    def edit_column(self):
        top = tkinter.Toplevel(self.root)
        entry_list = []
        for column in self.column:
            val_obj = tkinter.StringVar(top, value=column)
            entry = tkinter.Entry(top,textvariable=val_obj)
            entry_list.append(entry)
            entry.grid()
            
        def done():
            self.column = [entry.get() for entry in entry_list]
            self.tree.config(column=self.column)
            for column in self.column:
                self.tree.heading(column, text=column)  # 设置表头
            self.tree.update()
            top.destroy()
            
        button = tkinter.Button(top, text='完成', command=done)
        button.grid()
        top.mainloop()
    
    def edit_column_content(self, event):
        res = self.tree.identify('region', event.x, event.y)
        print(res)
        if res == 'heading':
            r = self.tree.identify('column', event.x, event.y)
            column_index = int(re.findall('^#(\d+)', r)[0]) - 1
            val = self.column[column_index]
            entry = tkinter.Entry(self.root)
            entry.insert('end', val)
            entry.xview('end')
            entry.focus()
            entry.place(x=event.x, y=event.y)
            
            def update_column_head(entry, button=None):
                print('update_column_head')
                val = entry.get().strip()
                self.column[column_index] = val
                self.column = self.column[:]
                self.tree.config(column=self.column)
                for column in self.column:
                    self.tree.heading(column, text=column)  # 设置表头
                self.tree.update()
                entry.destroy()
                
            entry.bind('<FocusOut>', lambda event: update_column_head(entry))
        elif res == 'cell':
            item = self.tree.identify('row', event.x, event.y)
            r = self.tree.identify('column', event.x, event.y)
            column_index = int(re.findall('^#(\d+)', r)[0]) - 1
            val = self.tree.item(item, 'values')
            entry = tkinter.Entry(self.root)
            entry.insert('end', val[column_index])
            entry.xview('end')
            entry.focus()
            
            def update_column_content(entry, button=None):
                print('update_column_content')
                val = entry.get().strip()
                self.tree.set(item, r, value=val)
                entry.destroy()

            x, y, w, h = self.tree.bbox(item, column=r)
            print(x, y, w, h)
            entry.place(x=x, y=y)
            entry.bind('<FocusOut>', lambda event: update_column_content(entry))
        
    def make_testcase(self):
        head = '\t'.join(self.column)
        head += '\n'
        content = ''
        for item in self.tree.get_children():
            values = self.tree.item(item, 'values')
            values = [t.strip() for t in values]
            content += '\t'.join(values)
            content += '\n'
        content = head + content
        with open('testcase.xlsx', 'w', encoding='utf-8') as fd:
            fd.write(content)
        try:
            with open('output.xlsx', 'wb') as fd:
                p = subprocess.Popen('allpairs.exe testcase.xlsx', stderr=subprocess.PIPE, stdout=fd)
        except Exception as e:
            messagebox.showerror('保存失败', e, parent=self.root)
            return
        else:
            err = p.stderr.read().decode('utf-8')
            if err:
                messagebox.showerror('保存失败', err, parent=self.root)
                return
            messagebox.showinfo('保存成功', '结果保存到文件output.xlsx', parent=self.root)
    
    def make_all_permutations(self):
        all_values = [[] for i in range(len(self.column))]
        for item in self.tree.get_children():
            values = self.tree.item(item, 'values')
            i = 0
            while i < len(values):
                if values[i].strip() != '': 
                    all_values[i].append(values[i]) 
                i += 1
        tmp = []
        ans = []
        
        def dfs(index):
            if index == len(self.column):
                ans.append(tmp[:])
                return
            for i in range(len(all_values[index])):
                tmp.append(all_values[index][i])
                dfs(index+1)
                tmp.pop()
                
        dfs(0)
        print('\t'.join(self.column))
        for i in ans:
            print('\t'.join(i))
        with open('all_permutations.csv', 'w+', encoding='utf-8') as fd:
            fd.write(','.join(self.column)+'\n')
            for i in ans:
                fd.write(','.join(i)+'\n')
        
    def pop_menu(self, event):
        res = self.tree.identify('region', event.x, event.y)
        menu = tkinter.Menu(self.root)
        if res == 'heading':
            r = self.tree.identify('column', event.x, event.y)
            column_index = int(re.findall('^#(\d+)', r)[0]) - 1
            val = self.column[column_index]
            menu.add_command(label=f'删除列: [{val}]', command=lambda: self.delete_column(val))
        elif res == 'cell':
            r = self.tree.identify('row', event.x, event.y)
            menu.add_command(label='删除该行', command=lambda: self.delete_row(r))
        menu.post(event.x_root, event.y_root)
    
    def add_row(self):
        vals = [DEFAULT_VALUE for i in range(len(self.column))]
        self.tree.insert('', 'end', values=vals, tags=('self',))
    
    def delete_row(self, item):
        self.tree.delete(item)
    
    def delete_column(self, column_name):
        # return
        column_index = self.column.index(column_name)
        all_values = []
        for item in self.tree.get_children():
            values = list(self.tree.item(item, 'values'))
            del values[column_index]
            all_values.append(values)
            self.tree.delete(item)
        
        self.column.remove(column_name)
        self.tree.config(column=self.column)
        for column in self.column:
            self.tree.heading(column, text=column)  # 设置表头
        for values in all_values:
            self.tree.insert('', 'end', values=values, tags=('self',))
        self.tree.update()
    
    def add_column(self):
        new_column_name = f'新增列{self.index}'
        self.column.append(new_column_name)
        self.index += 1
        self.tree.config(column=self.column)
        for column in self.column:
            self.tree.heading(column, text=column)  # 设置表头
        
        for item in self.tree.get_children():
            self.tree.set(item, new_column_name, value=DEFAULT_VALUE)
        
        self.tree.update()
        
        
if __name__ == '__main__':
    Gui().main()
