from tkinter import *

## GUI
window = Tk()

title = Label(window, text="Coast Water Quote Generator")
subtitle = Label(window, text="hella sick")
entry_1 = Entry(window)

title.grid()
subtitle.grid(row=1)

entry_1.grid(row=1,column=1)

## APP

def printName(event):
    print("Yo dawg")

button = Button(window, text="Print")
button.bind("<Button-1>",printName)
button.grid(row=2)







window.mainloop()