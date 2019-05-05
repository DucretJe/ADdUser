# coding: utf-8

from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter import messagebox

class Interface(Frame):
    """Main interface"""

    def __init__(self,window,**kwargs):

        Frame.__init__(self, window, **kwargs)
        self.pack(fill=BOTH)

        self.var_input = IntVar()
        self.waiting = IntVar()

    def validate(self):
        """ Method 'validate' when the user clic on the validate button """

        self.waiting.set(1)
        self.var = self.var_input.get()
        return self.var

    def simple_message(self, window, message):
        """ Display a simple message on the interface.
        Await a str in argument containing the message to display. """

        self.message = Label(window, text= message).pack(side=TOP)
        self.var_input = IntVar()
        self.button_quit = Button(window, text='Quit', command=self.quit).pack(side=LEFT)
        self.button_next = Button(window, text='Next', command=self.validate).pack(side=RIGHT)

    def set_filename(self):
        """ method for browsing files """
        FILETYPES = [ ("Excel Files", "*.xlsx") ]
        filename = StringVar(self)
        filename.set(askopenfilename(filetypes=FILETYPES))
        self.file = filename.get()
        return self.file

    def browse_message(self,window,message):
        """ Display message and creates the button in order to browse files """

        self.message = Label(window, text= message).pack(side=TOP)
        self.button_quit = Button(window, text='Quit', command=self.quit).pack(side=LEFT)
        self.button_browse = Button(window, text='browse', command=self.set_filename).pack(side=LEFT)
        self.button_next = Button(window, text='Next', command=self.validate).pack(side=RIGHT)

    def input_message(self,window,message):
        """ Display a message and let the user write something """

        self.message = Label(window, text= message).pack(side=TOP)
        self.var_input = StringVar()
        self.ligne_texte = Entry(window, textvariable=self.var_input, width=30).pack()
        self.button_quit = Button(window, text='Quit', command=self.quit).pack(side=LEFT)
        self.button_next = Button(window, text='Next', command=self.validate).pack(side=RIGHT)
