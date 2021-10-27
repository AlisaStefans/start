import os
import openpyxl
import shutil
from tkinter import *
from tkinter import messagebox as mb

root = Tk()

def destPhoto(source, destination, articl_own):
    try:
        source_png = (source + '.png')
        dest = shutil.move(source_png, destination)
    except:
        try:
            source_jpg = (str(source) + '.jpg')
            dest = shutil.move(source_jpg, destination)
        except:
            try:
                source_bmp = (str(source) + '.bmp')
                print('there')
                dest = shutil.move(source_bmp, destination)
            except:
                lbox_er.insert(END, str(str(articl_own)+ ' неизвестный формат или файл не найден'))
                raise ValueError('неизвестный формат')
        

def movePhoto():
    try:  
        wb = openpyxl.reader.excel.load_workbook(filename = "names.xlsx")
        sheet = wb.active
    except:
        mb.showinfo(title='Ошибка', message='Что-то не так с файлом names.xlsx')
        root.destroy()

    i = 2
    a = 2

    while a != None :
        try:
    
            articl_own = (sheet['B' + str(i)].value)
            articl_WB = (sheet['A' + str(i)].value)
            a = articl_own

            source = str(os.getcwd()) + "\\" + str(articl_own)
            destination = str(os.getcwd()) + "\\" + str(articl_WB) + "\\photo\\"

            if a is None:
                print('Готово')
                title_result.config(text='Результат:  Все сделано, начальника', bg='#66ff00')
                break
            else:
                if os.path.exists(str(os.getcwd()) + '/'+str(articl_WB)+'/photo/'):
                    print('здеся в проверке создания папки')
                    if os.path.exists(str(os.getcwd()) + '/'+str(articl_WB)+
                                    '/photo/' +
                                    str(articl_own) +
                                    '.png') or os.path.exists(str(os.getcwd())
                                                              + str(articl_WB)+ '/photo/' + str(articl_own)+
                                                              '.jpg') or os.path.exists(str(os.getcwd()) +
                                                                                        +str(articl_WB)+'/photo/' + str(articl_own)+ '.bmp'):
                        print('здеся в проверке создания файла')
                    else:
                        destPhoto(source, destination, articl_own)
                else:
                    os.makedirs(str(articl_WB) + "/photo/")
                    print('great dirrectory')
                    destPhoto(source, destination, articl_own)

        except:
            lbox_er.insert(END, str(articl_WB))
            try:
                os.rmdir(str(os.getcwd()) + "/" + str(articl_WB) + '/photo/')
                os.rmdir(str(os.getcwd()) + "/" + str(articl_WB))
                
            except:
                try:
                    wrong_er = str(articl_WB)
                    inx_wrong_er = (lbox_er.get(0, END).index(wrong_er))
                    lbox_er.delete(inx_wrong_er)
                    inx_wrong_er2 = (lbox_er.get(0, END).index(str(articl_own) + ' неизвестный формат или файл не найден'))
                    lbox_er.delete(inx_wrong_er2)
                except:
                    print('OOPS')
                
        finally:
            i = i + 1




def renamePhoto():
    #list_errors = []
    try:  
        wb = openpyxl.reader.excel.load_workbook(filename = "names.xlsx")
        sheet = wb.active
    except:
        mb.showinfo(title='Ошибка', message='Что-то не так с файлом names.xlsx')
        root.destroy()

    a = 2
    i = 2
    while a != None :
        try:
            articl_own = (sheet['A' + str(i)].value)
            articl_WB = (sheet['B' + str(i)].value)
            a = articl_own 
                
        #os.makedirs(str(articl) + "/photo")
        #source = str(os.getcwd()) + "/" + str(fotoname)
        #destination = str(os.getcwd()) + "/" + str(articl) + "/photo/"
        #dest = shutil.move(source, destination)
        #print(str(i) + str(dest))
            if a is None:
                print('Готово')
                title_result.config(text='Результат:  Все сделано, насяльника', bg='#66ff00')
                break
            else:
                os.rename(str(articl_own), str(articl_WB))
            
        except:
            print("ошибка с этим артикулом " + str(articl_own) + " номер строки " + str(i))
            #list_errors.append(str(articl_own))
            lbox_er.insert(END, str(articl_own))

        finally:
            i = i + 1
            continue
        

        #messagebox.showinfo(title='ok', message='ok')

root['bg'] = '#fafafa'
root.title('Переименование картинок')
root.wm_attributes('-alpha', 1)
root.geometry('500x500')
#root.resizable(width=False, height=False)

#создаем канву для виджетов
canvas = Canvas(root, height=500, width=500) 
canvas.pack()

frame = Frame (root)
frame.place(relwidth=1, relheight=1)

#frame_text = Frame (root, bg='white')
#frame_text.place(relwidth=1, relheight=0.5)

title = Label(frame, text='Работаем в дирректории:', bg='#66ffff', font=15)
title.place(relx=0.05, rely=0.05)

title_dir = Label(frame, text= str(os.getcwd()), bg='#66ffff')
title_dir.place(relx=0.05, rely=0.15)

title_result = Label(frame, text='Результат:', bg='#66ffff')
title_result.place(relx=0.05, rely=0.45)

title_result2 = Label(frame, text='В этих артикулах произошла ошибка: ', bg='#66ffff')
title_result2.place(relx=0.05, rely=0.55)

lbox_er= Listbox(frame, width=70, selectmode=EXTENDED)
lbox_er.place(relx=0.05, rely=0.65)
scroll = Scrollbar(command = lbox_er.yview)
scroll.pack(side=LEFT, fill=Y)
lbox_er.config(yscrollcommand=scroll.set)


start_btn = Button(frame, bg='#66ffff', text='Переименовать картинки', font=30, command=renamePhoto)
start_btn.place(relx=0.05, rely=0.25)
start_btn2 = Button(frame, bg='#66ffff', text='Переместить картинки', font=30, command=movePhoto)
start_btn2.place(relx=0.05, rely=0.35)

root.mainloop()




    
