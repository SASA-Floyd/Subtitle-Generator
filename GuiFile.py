from tkinter import filedialog, Tk
root = Tk()
root.filename = filedialog.askopenfilename(
    initialdir="/",
    title="파일을 선택하세요",
    filetypes=(("Microsoft Excel Files", ".xlsx"),("all files", "*.*")))


print(root.filename)
