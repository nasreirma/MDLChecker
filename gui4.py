

from pathlib import Path
import win32com.client as win32

from tkinter import Tk, Canvas, Entry,  Button, PhotoImage , filedialog , Label , messagebox
from tkinter.ttk import Progressbar , Style

OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path(r"assets/frame0")


def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)


window = Tk()
style = Style()
style.theme_use('default')

window.geometry("400x500")
window.configure(bg = "#3A7FF6")
window.title("MDL FIle Checker")
logo = PhotoImage(file=ASSETS_PATH / "Logo.png")
window.call('wm', 'iconphoto', window._w, logo)


def select_xlsx():
    global xlsx_file
    xlsx_file = filedialog.askopenfilename( title="Select MDL file", filetypes=(("Excel files", "*.xlsx"),))
    entry_1.delete(0, "end")  # Clear the current value of the Entry widget
    entry_1.insert(0, xlsx_file)  # Set the value of the Entry widget to the selected file path


def select_txt():
    global txt_file
    txt_file = filedialog.askopenfilename( title="Select Variable file", filetypes=(("Text files", "*.txt"),))
    entry_2.delete(0, "end")  # Clear the current value of the Entry widget
    entry_2.insert(0, txt_file)  # Set the value of the Entry widget to the selected file path


def check_files():
    try:
        MDLFile = xlsx_file
    except:
        messagebox.showerror(
            title="Empty Fields!", message="Please enter MDL File.")
    try:
        VarFile = txt_file
    except:
        messagebox.showerror(
            title="Empty Fields!", message="Please enter Variable File.")

    progress['value']=0

    f = open(VarFile, "r").readlines()
    File = [i.strip("'\n") for i in f]

    # Opening Workbook and Worksheet
    xl = win32.Dispatch('Excel.Application')
    wb = xl.Workbooks.Open(MDLFile)
    ws = wb.Worksheets('Mapping_File_MDL_LINK')

    def rgbToInt(rgb):
        colorInt = rgb[0] + (rgb[1] * 256) + (rgb[2] * 256 * 256)
        return colorInt

    # Update Values
    total_rows = ws.UsedRange.Rows.Count + 2
    for i in range(1, total_rows):
        try:
            y = str((ws.Cells(i, 3).Value).strip())
            if y in File:
                ws.Cells(i, 5).Value = "OK"
                ws.Cells(i, 5).Interior.Color = rgbToInt((0, 255, 0))
            elif y not in File and y.startswith('Platform'):
                ws.Cells(i, 5).Value = "NOK"
                ws.Cells(i, 5).Interior.Color = rgbToInt((255, 0, 0))
        except:
            pass
        progress['value']=int((i / (total_rows-1))*100)
        label_1['text']=str(int((i / (total_rows-1))*100))+'%'
        window.update_idletasks()

    ws.Columns("E:E").AutoFilter(1)

    # Close and save the workbook
    wb.Close(True)
    xl.Quit()

    messagebox.showinfo(title="MDL File Checker",message="MDL File checking is finished.")



canvas = Canvas(
    window,
    bg = "#3A7FF6",
    height = 500,
    width = 400,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge"
)

canvas.place(x = 0, y = 0)
canvas.create_rectangle(
    0.0,
    0.0,
    400.0,
    500.0,
    fill="#FFFFFF",
    outline="")

entry_image_1 = PhotoImage(
    file=relative_to_assets("entry_1.png"))
entry_bg_1 = canvas.create_image(
    183.0,
    153.33333587646484,
    image=entry_image_1
)
entry_1 = Entry(
    bd=0,
    bg="#F1F1FF",
    fg="#000716",
    highlightthickness=0
)
entry_1.place(
    x=39.0,
    y=131.11111450195312,
    width=288.0,
    height=42.44444274902344
)

entry_image_2 = PhotoImage(
    file=relative_to_assets("entry_2.png"))
entry_bg_2 = canvas.create_image(
    183.0,
    249.99999237060547,
    image=entry_image_2
)
entry_2 = Entry(
    bd=0,
    bg="#F1F1FF",
    fg="#000716",
    highlightthickness=0
)
entry_2.place(
    x=39.0,
    y=227.77777099609375,
    width=288.0,
    height=42.44444274902344
)

button_image_1 = PhotoImage(
    file=relative_to_assets("button_1.png"))
button_1 = Button(
    image=button_image_1,
    borderwidth=0,
    highlightthickness=0,
    command=lambda: select_xlsx(),
    relief="flat"
)
button_1.place(
    x=342.0,
    y=131.11111450195312,
    width=37.0,
    height=44.44444274902344
)

button_image_2 = PhotoImage(
    file=relative_to_assets("button_2.png"))
button_2 = Button(
    image=button_image_2,
    borderwidth=0,
    highlightthickness=0,
    command=lambda: select_txt(),
    relief="flat"
)
button_2.place(
    x=342.0,
    y=227.77777099609375,
    width=37.0,
    height=44.44444274902344
)

button_image_3 = PhotoImage(
    file=relative_to_assets("button_3.png"))
button_3 = Button(
    image=button_image_3,
    borderwidth=0,
    highlightthickness=0,
    command=lambda: check_files(),
    relief="flat"
)
button_3.place(
    x=100.0,
    y=375.6666564941406,
    width=200.0,
    height=44.44444274902344
)

style.configure(
    "custom.Horizontal.TProgressbar",
    troughcolor='#F1F1FF',
    bordercolor='#F1F1FF',
    background='#0F3C93',
    thickness=20
)

progress = Progressbar(
    window,
    orient="horizontal",
    length=300,
    mode="determinate",
    style="custom.Horizontal.TProgressbar"
)

progress.place(
    x=35,
    y=320
)

canvas.create_text(
    41.0,
    102.22222137451172,
    anchor="nw",
    text="MDL File :",
    fill="#000000",
    font=("ABeeZee Regular", 14 * -1)
)

canvas.create_text(
    41.0,
    201.11111450195312,
    anchor="nw",
    text="Variable File :",
    fill="#000000",
    font=("ABeeZee Regular", 14 * -1)
)

label_1 = Label(
    anchor="nw",
    text="0%",
    background="#FFFFFF"
)


label_1.place(
    x=349.0,
    y=320.0
)

image_image_1 = PhotoImage(
    file=relative_to_assets("image_1.png"))
image_1 = canvas.create_image(
    199.0,
    49.33332824707031,
    image=image_image_1
)

image_image_2 = PhotoImage(
    file=relative_to_assets("image_2.png"))
image_2 = canvas.create_image(
    200.0,
    467.0,
    image=image_image_2
)
window.resizable(False, False)
window.mainloop()
