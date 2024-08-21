import sys
import os
import time
from tkinter import *
from tkinter import filedialog
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
import hybridity_logics as hybridity
import threading
from tkinter import ttk
import win32event
import win32api
import winerror
from PIL import ImageTk, Image
from class_hybridity import HybridQC

LOCK_FILE_PATH = "my_application.lock"
# import pyglet

class MainWindow():
    def __init__(self, geometry='900x615') -> None:
        self.mutex = win32event.CreateMutex(None, 1, "YourAppMutexName")
        if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
            # If another instance is already running, exit
            # print("Another instance of the application is already running.")
            return
        
        # if os.path.exists(LOCK_FILE_PATH):
        #     print("Another instance of the application is already running. Exiting...")
        #     return

        # # Create the lock file
        # self.create_lock_file()
        # self.bg = '#fffff7'
        self.loading_page()
        self.bg = '#FFFFF7'
        
        self.geometry = geometry
        self.showWindow()

    def task(self):
        time.sleep(3)
        self.root.destroy()    

    
    def center_window(self, window):
        window.update_idletasks()  # Ensure window dimensions are updated

        # Get screen dimensions
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()

        # Calculate window dimensions
        window_width = window.winfo_width()
        window_height = window.winfo_height()

        # Calculate position for the window to be centered
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        # Set the window's geometry to be centered
        window.geometry(f"+{x}+{y}")
        
        
    def loading_page(self):
        self.root = Tk()
        root = self.root
        root.overrideredirect(True)
        root.title("Loading...")
        # root.geometry("200x100")
        # icon = PhotoImage(file=resource_path('png.png'))
        root.resizable(False, False)
        image = ImageTk.PhotoImage(Image.open("assets/logo.jpg"))
        # Create a label with loading text
        label = Label(root, image=image)
        label.pack()


        # label.configure(background='systemTransparent')
        root.after(600, self.task)
        #Put the window in the middle
        self.center_window(root)
        root.mainloop()


    def showWindow(self):
        self.window = Tk(screenName='hybrid')
        # ImageTk.PhotoImage(Image.open("logo.jpg"))
        icon = ImageTk.PhotoImage(Image.open("assets/icon.png"))
        self.window.geometry(self.geometry)
        self.window.title('HybridQC')
        self.window.resizable(width=False, height=False)
        self.window.config(bg=self.bg)
        self.center_window(self.window)
        self.window.iconphoto(True,icon)

        self.toplabel = Label(self.window, text='HybridQC',
                              font=('Georgia', 32, 'bold'),
                              fg='black', bg=self.bg)
        self.toplabel.pack(pady=(30, 40))
        
        self.fileEntry = Entry(self.window,
                               font=('Arial', 16),
                               fg='black',
                               bg=self.bg, highlightthickness=2,
                               width=40, highlightbackground = "black", highlightcolor= "black")
        self.fileEntry.pack()
        
        self.info = Label(self.window,
                         text='Select xlsx file',
                         font=('Arial', 10, 'bold'),
                         fg='black', bg=self.bg)
        self.info.pack(pady=(0,5))
        
        self.selectFileButton = Button(font=('Arial', 12, 'bold'),
                        fg=self.bg,
                        bg='black',
                        width=10,
                        text='Select File',
                        relief=SOLID,
                        cursor='hand2',
                        command=self.selectFile)
        
        self.create_entry_row(self.window)
        self.update_button_position(None)

        # Bind the <Configure> event to the function
        self.fileEntry.bind('<Configure>', self.update_button_position)
        
        self.lastLabel(self.window)
        
        # self.progressBar(self.window).pack()
        # self.window
        # pb = ttk.Progressbar(
        #     self.window,
        #     orient='horizontal',
        #     mode='indeterminate',
        #     style="CustomProgressbar",
        #     length=100,
        # )
        
        # pb = ttk.Progressbar(
        #     self.window,
        #     orient='horizontal',
        #     mode='indeterminate',
        #     length=100,
        # )
        # pb.pack()
        s = ttk.Style()
        s.theme_use('alt')
        s.configure("red.Horizontal.TProgressbar", background='red')
        s.configure("yellow.Horizontal.TProgressbar", background='yellow')
        s.configure("green.Horizontal.TProgressbar", background='green')
        s.configure("black.Horizontal.TProgressbar", background='black')
        
        self.pb = ttk.Progressbar(
            self.window,
            orient='horizontal',
            mode='indeterminate',
            length=100,
            # style="yellow.Horizontal.TProgressbar"  # Specify the custom style
        )
        self.pb.config(style='black.Horizontal.TProgressbar')
        # Call the function to change appearance
        # self.change_progressbar_appearance(progressbar)
        self.pb.pack(pady=3)

        # progressbar.start()
        watermark = Label(text="© AJIBADE Y.A & Ongom P.O, IITA, Nigeria. 2024", font=("Times New Roman",9,'bold'), bg=self.bg)
        watermark.place(relx=1.0, rely=1.0, x=0, y=0, anchor=SE)
        self.window.mainloop()
        self.remove_lock_file()


    def create_lock_file(self):
        # Create the lock file
        with open(LOCK_FILE_PATH, "w") as lock_file:
            lock_file.write("")

    def remove_lock_file(self):
        # Remove the lock file
        if os.path.exists(LOCK_FILE_PATH):
            os.remove(LOCK_FILE_PATH)
            
               
    def lastLabel(self, window):
        self.log = Label(window,
                         text='',
                         font=('Arial', 14,),
                         fg='black', bg=self.bg)
        self.log.pack(pady=(20, 0))
    
    
    def change_progressbar_appearance(self, progressbar):
        style = ttk.Style()

        # Change height
        # style.configure("TProgressbar", thickness=10)

        # Change color
        style.configure("yellow.Horizontal.TProgressbar", troughcolor="black", background="black", thickness=50)

        progressbar.style = style
        
        
    # def progressBar(self, window):
    #     pb = ttk.Progressbar(
    #         window,
    #         orient='horizontal',
    #         mode='indeterminate',
    #         style="TProgressbar",
    #         length=100,
    #     )
        
    #     return pb
        
    
    def selectFile(self):
        self.fileEntry.config(state='normal')
        self.fileEntry.delete(0, END)
        filepath = filedialog.askopenfile(title="select the excel file in xlsx format")
        ourfile = filepath.name
        self.fileEntry.insert(0, ourfile)
        self.fileEntry.xview_moveto(1.0)
        # self.fileEntry.icursor(END)
        self.fileEntry.config(state='readonly')
        
 
    def validateInput(self, perc, threshold):
        try:
            perc = int(perc)
        except ValueError:
            perc = threshold
        return perc
    
    def dotheJob(self):
        window = self.window
        try:
            self.runButton.config(state=DISABLED)
            self.log.config(text='Working...', fg='black')
            
            perc_poly = self.perc_polymorphic_threshold.get()
            perc_missing = self.perc_missing_threshold.get()
            perc_hybrid = self.perc_hybridity_threshold.get()
            
            perc_poly = self.validateInput(perc_poly, 20)
            perc_missing = self.validateInput(perc_missing, 20)
            perc_hybrid = self.validateInput(perc_hybrid, 20)

            savepath = self.saveas()
            if savepath is not None:
                hybrid = HybridQC(filename = self.fileEntry.get())
                
                hybridity_thread = threading.Thread(
                    target=hybridity.hybridity,
                    args=[self.fileEntry.get(), savepath, self.log, int(perc_missing), int(perc_poly), int(perc_hybrid)]
                )

                self.pb.start()  # Start the progress bar before starting the thread
                hybridity_thread.start()

                # Periodically check if the thread is alive and update the GUI
                while hybridity_thread.is_alive():
                    window.update()  # Process Tkinter events
                    # time.sleep(0.1)  # Avoid excessive CPU usage

                self.log.config(text='COMPLETED', fg='#008000', font=('Arial', 14, 'bold'))
                # self.citationWindow()
        except PermissionError:
            self.log.config(text="Your output file should not be opened in another app", fg="#800000")
        except Exception as e:  # Catch any other unexpected errors
            self.log.config(text=f"An error occurred", fg="#800000")
        finally:
            self.pb.stop()  # Stop the progress bar
            self.runButton.config(state=NORMAL)


    def saveas(self) -> str:
        filepath = filedialog.asksaveasfile(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if filepath:
            return str(filepath.name)
        else:
            self.log.config(text="Canceled! Please select destination to continue", fg="#800000")
            return None
        
    def update_button_position(self, event):
        x_position = self.fileEntry.winfo_x() + self.fileEntry.winfo_width() + 32
        y_position = self.fileEntry.winfo_y()
        self.selectFileButton.place(x=x_position, y=y_position)
        
        x_position_entry = self.fileEntry.winfo_x()
        y_position_entry = self.fileEntry.winfo_y() + self.fileEntry.winfo_height()
        
        self.info.place(x=x_position_entry, y=y_position_entry)
    
    
    def defaultThreshold(self):
        is_checked = self.var.get()
        if is_checked:
            self.perc_hybridity_threshold.delete(0, END)
            self.perc_missing_threshold.delete(0, END)
            self.perc_polymorphic_threshold.delete(0, END)
            
            self.perc_hybridity_threshold.insert(0, '50')
            self.perc_missing_threshold.insert(0, '20')
            self.perc_polymorphic_threshold.insert(0, '20')
            
            self.perc_hybridity_threshold.config(state='readonly')
            self.perc_missing_threshold.config(state='readonly')
            self.perc_polymorphic_threshold.config(state='readonly')
            
        else:
            self.perc_hybridity_threshold.config(state='normal')
            self.perc_missing_threshold.config(state='normal')
            self.perc_polymorphic_threshold.config(state='normal')
    
    
    def create_entry_row(self, window):
        bg = '#FFFFCC'
        '#fffc9d'
        row_frame = Frame(window, bg=bg, bd=2, relief=SOLID,
                          width=400, highlightbackground="#fffc9d",
                          height=353)  # Create a frame for the row
        #ffffcc
        Label(row_frame, font=('Arial', 18, 'bold'),
              text="Input Thresholds",
              bg=bg).grid(row=0, column=0, columnspan=3, pady=(10, 20))

        Label(row_frame, font=('Arial', 16,),
              text="Min % polymorphic:", bg=bg).grid(row=1, column=0, padx=(35, 0))
        self.perc_polymorphic_threshold = Entry(row_frame, font=('Arial', 15),
                                                bg=self.bg, highlightthickness=1,
                                                highlightbackground = "black", highlightcolor= "black",
                                                fg='black', width=5)
        self.perc_polymorphic_threshold.grid(row=1, column=1, padx=20, pady=10)

        Label(row_frame, font=('Arial', 16,),
              text="Max % missing:", bg=bg).grid(row=2, column=0)
        self.perc_missing_threshold = Entry(row_frame, font=('Arial', 15), 
                                            bg=self.bg, highlightthickness=1,
                                                highlightbackground = "black", 
                                                highlightcolor= "black",
                                            fg='black', width=5)
        self.perc_missing_threshold.grid(row=2, column=1, padx=20, pady=10)

        Label(row_frame, font=('Arial', 16,),
              text="Min % hybridity:", bg=bg).grid(row=3, column=0)
        self.perc_hybridity_threshold = Entry(row_frame, font=('Arial', 15), 
                                              bg=self.bg, highlightthickness=1,
                                                highlightbackground = "black",
                                                highlightcolor= "black",
                                                fg='black', width=5)
        self.perc_hybridity_threshold.grid(row=3, column=1, padx=20, pady=10)

        self.var = BooleanVar()
        self.defaultThresholdBtn = Checkbutton(row_frame, text="Use default threshold",
                                    variable=self.var, bg=bg,
                                    command=self.defaultThreshold)
        self.defaultThresholdBtn.grid(row=4, column=0, pady=10)
        
    
        self.runButton = Button(row_frame, font=('Arial', 15, 'bold'),
                   bg='black', width=18, height= 1,
                   text='Run', fg=self.bg,
                #    pady=(15, 5),
                   relief=SOLID,
                   cursor='hand2',
                   command=self.dotheJob)
        self.runButton.grid(row=5, column=0, columnspan=3, pady=(10, 20))
        
        row_frame.pack(pady=(40, 0))  # Pack the entire row frame
        
    def citationWindow(self):
        window = self.window
        popup = Toplevel(window)
        popup.attributes('-topmost', True)
        popup.resizable(width=False, height=False)
        # Set the size of the pop-up window
        popup_width = 500
        popup_height = 100


        # Calculate the position to center the pop-up window relative to the main window
        x_position = window.winfo_x() + (window.winfo_width() - popup_width) // 2
        y_position = window.winfo_y() + (window.winfo_height() - popup_height) // 2

        # Set the geometry of the pop-up window
        popup.geometry(f"{popup_width}x{popup_height}+{x_position}+{y_position}")
        
        
        reference = Label(popup, text='Add the citation to your Reference', font=('Arial' ,12, 'bold'))
        reference.pack(pady=(0, 2))
        
        reference_entry = Entry(popup, width=400,
                                borderwidth=0,
                                highlightthickness=0, font=(14))
        
        reference_entry.insert(0, "      Patrick O., Yakub A.")
        reference_entry.config(state='readonly')
        reference_entry.pack()
        

        def close():
            get_referrence = reference_entry.get()
            popup.destroy()
            
        closeBtn = Button(popup, text='CLOSE', command=close)
        closeBtn.pack(pady=(20,0))
            
app = MainWindow()
# a.showWindow()
# a.update_button_position(None)
# a.fileEntry.bind('<Configure>', a.update_button_position)

