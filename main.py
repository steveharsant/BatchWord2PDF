import tkinter
import customtkinter
import os
import glob
import sys
from docx2pdf import convert
import winreg


version = '0.2.0'

# Redirects output of docx2pdf which resolves an issue where the library doesn't run when compiled. See:
# https://stackoverflow.com/questions/74787311/error-with-docx2pdf-after-compiling-using-pyinstaller
sys.stderr = open('consoleoutput.log', 'w')

### * functions *###


def pick_folder():
    folder_path = tkinter.filedialog.askdirectory()
    folder_path_textbox.delete('0.0', tkinter.END)
    folder_path_textbox.insert('0.0', folder_path)
    message_label.configure(text='')


def start_conversion():
    convert_button.configure(text='Converting...')
    app.update()
    folder_path = (folder_path_textbox.get(
        '0.0', tkinter.END)).strip().replace('\\', '/')

    if not os.path.exists(folder_path):
        message_label.configure(text='The folder does not exist')

    else:
        doc_files = glob.glob(folder_path + "/*.docx") + \
            glob.glob(folder_path + "/*.doc")

        i = 1
        for file in doc_files:
            message_label.configure(
                text='Processing file {} of {}'.format(i, len(doc_files)))
            app.update()
            convert(file)
            i += 1

        message_label.configure(text='Conversion complete')

    convert_button.configure(text='Convert')


### * GUI *###
customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("blue")

app = customtkinter.CTk()
app.geometry("650x180")
app.title('Batch Word2PDF Converter')
app.resizable(False, False)

usage_label = customtkinter.CTkLabel(
    master=app, text="Select a folder with Word documents, then click 'Convert'")
usage_label.place(relx=0.5, rely=0.15, anchor=tkinter.CENTER)

folder_path_textbox = customtkinter.CTkTextbox(
    app, height=20, width=600, activate_scrollbars=False)
folder_path_textbox.place(relx=0.5, rely=0.4, anchor=tkinter.CENTER)

picker_button = customtkinter.CTkButton(
    master=app, text="📁", command=pick_folder, width=50)
picker_button.place(relx=0.925, rely=0.4, anchor=tkinter.CENTER)

convert_button = customtkinter.CTkButton(
    master=app, text="Convert", command=start_conversion)
convert_button.place(relx=0.5, rely=0.65, anchor=tkinter.CENTER)

message_label = customtkinter.CTkLabel(master=app, text='')
message_label.place(relx=0.5, rely=0.85, anchor=tkinter.CENTER)

app.mainloop()
