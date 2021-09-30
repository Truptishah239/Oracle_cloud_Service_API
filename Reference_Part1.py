
import pandas as pd
import tkinter as tk
from tkinter import simpledialog
from tkinter import Frame, Tk, Button
import tkinter.messagebox as msg

class MyDialog3(tk.simpledialog.Dialog):
    def __init__(self, parent):
        self.reference_ID = None
        super().__init__(parent)

    def body(self, Frame):
        # print(type(Frame)) # tkinter.Frame
        self.reference_ID_label = tk.Label(Frame, width=16, text="reference_ID",fg="white",bg="#eb67ad")
        self.reference_ID_label.pack()
        self.reference_ID_box = tk.Entry(Frame, width=25,highlightthickness=2)
        self.reference_ID_box.config(highlightbackground = "#f0c7dd", highlightcolor= "#8a6679")
        self.reference_ID_box.pack()
        return Frame

    def ok_pressed(self):
        # print("ok")
        self.reference_ID = self.reference_ID_box.get()
        self.destroy()

    def cancel_pressed(self):
        # print("cancel")
        self.destroy()


    def buttonbox(self):
        self.ok_button = tk.Button(self, text='OK', width=5, bg='#eb67ad',command=self.ok_pressed)
        self.ok_button.pack(side="left")
        cancel_button = tk.Button(self, text='Cancel', width=5,bg='#eb67ad', command=self.cancel_pressed)
        cancel_button.pack(side="right")
        self.bind("<Return>", lambda event: self.ok_pressed())
        self.bind("<Escape>", lambda event: self.cancel_pressed())    
