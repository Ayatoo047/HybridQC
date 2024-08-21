import sys
import os
import time
from tkinter import *
from tkinter import filedialog
#sys.path.append(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(os.path.dirname(sys.executable))
import hybridity_logics as hybridity
import threading
from tkinter import ttk
from PIL import ImageTk, Image
LOCK_FILE_PATH = "my_application.lock"
# import pyglet


import time

    
def calc_perc_polymorphic(polymorphic, no_markers):
    return (int(polymorphic)/int(no_markers)) * 100

def calc_perc_missing(missing, polymorphic):
    return (int(missing)/int(polymorphic)) * 100

def perc_hybridity(true, polymorphic, missing):
    return (int(true))/(int(str(polymorphic))-int(str(missing)))* 100
    
    
# def hybridity(filename, saveas, log, min_missing_percentage=20, min_perc_polymorphic=20, min_perc_hybridity=50):
#     # start = time.time()
#     # try:
#     wb = load_workbook(filename, data_only=True)
#     # except Exception as e:
#     #     log.config(text=f"No such file or directory", fg="#800000")
#     #     raise
    
    
#     sheet = wb.worksheets[0]
#     sheet.title = 'Results'
    
#     toskip = []
#     marker_success = []
#     # min_missing_percentage = 20
#     # min_perc_polymorphic = 0

#     sheet.insert_cols(3, 7)
#     # sheet.insert_cols(7, 3)
    
#     undefined_missing_data = 0
#     undefined_unpolymorphic_parent = 0
    
#     parent_polymophic = 0
#     percent_polymorph = 0
    
#     TRUE = 0
#     FAILED = 0

#     no_markers = sheet.max_column - 10

#     green = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
#     red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
#     orange = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
#     grey = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")

#     polymophic_col = 2
#     perc_polymorphic = 3
#     trueF1 = 4
#     missing_col = 5
#     perc_missing = 6
#     hybridity_col = 7
#     status_col = 8
#     parent_column = 9
#     marker_start = 10
    
#     def color_previous(col_number):
#         sheet[f'{value[col_number].column_letter}{value[col_number].row - 1}'].fill = grey

    
#     track_parent = 0
#     track_parent_on_F1 = 0
#     for k, (value, value2) in enumerate(zip(sheet.iter_rows(min_row=2, max_row=sheet.max_row),
#                                                 sheet.iter_rows(min_row=3, max_row=sheet.max_row)), 1):
#         if value[parent_column].value == "Parent" and value2[parent_column].value == 'Parent':
#             track_parent+=1
#             parent_polymophic = 0
#             value[parent_column].fill = green
#             value2[parent_column].fill = green
#             for i in range(marker_start, sheet.max_column):
#                 z = i - marker_start
#                 if len(marker_success) < no_markers:
#                     marker_success.append([z, value[i].column_letter, 0])

#                 if value[i].value != value2[i].value and all(v.value not in ['Uncallable', '?'] for v in [value[i], value2[i]]):
#                     marker_success[z] = ([z, value[i].column_letter, ((marker_success[z])[2] + 1)])
                    
#                     parent_polymophic += 1

#                     value[i].fill = green
#                     value2[i].fill = green
#                 else:

#                     toskip.append((track_parent, value[i].column_letter))

#                     value[i].fill = red
#                     value2[i].fill = red
#                 # progressBar(k, sheet.max_row, pb, window, log)

#             value[polymophic_col].value = parent_polymophic
#             value[polymophic_col].fill = grey
            
#             value2[polymophic_col].value = value[polymophic_col].value
#             value2[polymophic_col].fill = grey
            
#             percent_polymorph = calc_perc_polymorphic(value[polymophic_col].value, no_markers)
            
#             value[perc_polymorphic].value = percent_polymorph
#             value[perc_polymorphic].fill = grey
            
#             value2[perc_polymorphic].value = value[perc_polymorphic].value
#             value2[perc_polymorphic].fill = grey

#     for k, (value, value2) in enumerate(zip(sheet.iter_rows(min_row=2, max_row=sheet.max_row),
#                                                 sheet.iter_rows(min_row=3, max_row=sheet.max_row)), 1):
#         if (value[parent_column].value == 'Parent' and value2[parent_column].value == 'F1') or (value[parent_column].value == 'F1' and value2[parent_column].value == 'F1'):
            
#             # value2[polymophic_col].value = value[polymophic_col].value
#             value2[polymophic_col].fill = grey
            
#             # polymophic = value2[polymophic_col].value
            
#             missiing = 0
#             true = 0
#             if (value[parent_column].value == 'Parent' and value2[parent_column].value == 'F1'):
                
#                 percent_polymorph = int(value[perc_polymorphic].value)
#                 parent_polymophic = int(value[polymophic_col].value)
#                 track_parent_on_F1 += 1

#             for i in range(marker_start, sheet.max_column):
#                 skip = (track_parent_on_F1, value[i].column_letter)
            
#                 if value2[i].value in ['Uncallable', '?'] and skip not in toskip:
#                     missiing += 1
#                 if value2[i].value not in ['G:G', 'A:A', 'C:C', 'T:T', 'Uncallable', '?'] and skip not in toskip:
#                     value2[i].fill = orange
#                     true += 1
                

#             value2[missing_col].value = missiing
#             value2[trueF1].value = true
            
#             # color 
#             value2[missing_col].fill = grey
#             value2[trueF1].fill = grey
            
            
#             if value[missing_col].value == None:
#                 value[missing_col].fill = grey
                
#                 color_previous(missing_col)
#                 # sheet[f'{value[missing_col].column_letter}{value[missing_col].row - 1}'].fill = grey
                
                
#             if value[trueF1].value == None:
#                 value[trueF1].fill = grey
                
#                 color_previous(trueF1)
            
#             # color grey  
                
            
#             try:
#                 percent_missing = calc_perc_missing(missiing, parent_polymophic)
#             except ZeroDivisionError:
#                 percent_missing = 999
            
#             # value2[perc_polymorphic].value = percent_polymorph
#             #color
#             value2[perc_polymorphic].fill = grey
            
#             if value[perc_polymorphic].value == None:
#                 value[perc_polymorphic].fill = grey
#                 color_previous(perc_polymorphic)
            
#             if percent_missing != 999:
#                 value2[perc_missing].value = percent_missing
#             else:
#                 value2[perc_missing].value = 'NA'
                
#             #color
#             value2[perc_missing].fill = grey
            
#             if value[perc_missing].value == None:
#                 value[perc_missing].fill = grey
#                 color_previous(perc_missing)
                
#             if (percent_polymorph <= min_perc_polymorphic):
#                 value2[hybridity_col].value = 'NA'
#                 value2[status_col].value = 'Undetermine: Parent not polymorphic'
#                 undefined_unpolymorphic_parent += 1
            
#             if (((percent_missing >= min_missing_percentage) or (percent_missing == 999)) and value2[status_col].value is None):
#                 value2[hybridity_col].value = 'NA'
#                 value2[status_col].value = 'Undetermine: missing data'
#                 undefined_missing_data += 1
            
#             value2[hybridity_col].fill = grey
#             value2[status_col].fill = grey
            
            
#             if value[hybridity_col].value == None:
#                 value[hybridity_col].fill = grey
#                 color_previous(hybridity_col)
#             if value[status_col].value == None:
#                 value[status_col].fill = grey
#                 color_previous(status_col)
#             try:
#                 if value2[hybridity_col].value is None:
#                     percent_hybridity = perc_hybridity(true, parent_polymophic, missiing)
                    
#                     if percent_hybridity <= min_perc_hybridity and value2[status_col].value is None:
#                         value2[status_col].value = 'FAILED'
#                         FAILED += 1
#                     else:
#                         value2[status_col].value = 'TRUE'
#                         TRUE += 1
                

#                     value2[hybridity_col].value = percent_hybridity
#                     value2[hybridity_col].fill = grey
                    
                    
#                     if value[hybridity_col].value == None:
#                         value[hybridity_col].fill = grey
#                         color_previous(hybridity_col)
                    
#                     # value2[hybridity_col].value = (int(str(value2[3].value))/(int(str(value2[polymophic_col].value))-int(str(value2[missing].value)))) * 100
#                 # if int(value2[hybridity_col].value) > 100:
#                 #     print(value2[hybridity_col].coordinate)
              
#             except ValueError:
#                 pass
#             except ZeroDivisionError:
#                 value2[hybridity_col].value = 'NA'
#                 value2[hybridity_col].fill = grey
#                 if value[hybridity_col].value == None:
#                     value[hybridity_col].fill = grey
#                     color_previous(hybridity_col)
#         else:
#             continue 
        

#     sheet['H1'].value = '%hybridity'
#     sheet['D1'].value = '%polymorphic'
#     sheet['G1'].value = '%missing'
#     sheet['I1'].value = 'Status'
#     sheet['C1'].value = '#polymorphic'
#     sheet['E1'].value = '#true'
#     sheet['F1'].value = '#missing'
    
    
#     sheet['C1'].fill = grey
#     sheet['E1'].fill = grey
#     sheet['F1'].fill = grey
#     sheet['H1'].fill = grey
#     sheet['D1'].fill = grey
#     sheet['G1'].fill = grey
#     sheet['I1'].fill = grey
#     # wb.save('compiled_hybridity.xlsx')
#     # print(TRUE, FAILED, undefined_missing_data, undefined_unpolymorphic_parent)
#     TOTAL = TRUE + FAILED + undefined_missing_data + undefined_unpolymorphic_parent
    
    
#     data = [
#         ["Status", 'Value'],
#         ["TRUE", TRUE],
#         ["FAILED", FAILED],
#         ["Undetermined: Parent not polymorphic", undefined_unpolymorphic_parent],
#         ["Undetermined: Missing data", undefined_missing_data],
#         ["TOTAL", TOTAL],
#     ]

#     chartSheet = wb.create_sheet(title='Hybridity')
#     for row in data:
#         chartSheet.append(row)
    
#     pie_chart = PieChart()
#     data = Reference(chartSheet, min_col=2, min_row=1, max_col=2, max_row=len(data)-1)
#     categories = Reference(chartSheet, min_col=1, min_row=2, max_row=len(data))

#     pie_chart.add_data(data, titles_from_data=True)
#     pie_chart.set_categories(categories)
#     pie_chart.title = "HYBRIDITY PIE CHART"
    
#     # Add data labels to the chart
#     pie_chart.dataLabels = label.DataLabelList()
#     pie_chart.dataLabels.showPercent = True  # Display percentage
#     # pie_chart.dataLabels.showVal = True  # Display actual values
    
#     pie_chart.dataLabels.position = "bestFit"  # Position labels at best places
#     pie_chart.dataLabels.number_format = "0.0%"  # Format percentages

    
#     chartSheet.add_chart(pie_chart, "H8")
    
    
#     barData = [
#             #    ["SNPs", 'Total Parents', "No of Polymorphic Parent", 'Frequenxy of Success'],
#                [sheet[f'{i[1]}1'].value, track_parent, i[2], ((i[2]/track_parent)*100)] for i in marker_success
#                ]
#     barData.insert(0, ["SNPs", '#Parent Combination', "Polymorphism Frequency", 'Marker Efficiency (%)'])
#     barSheet = wb.create_sheet(title='Efficiency')
    
#     for row in barData:
#         barSheet.append(row)
        
        
#     barchart = BarChart()
#     barchart.title = "SNPs Performance"
#     barchart.x_axis.title = "Markers"
#     barchart.y_axis.title = "Marker Efficiency"
    
#     data = Reference(barSheet, min_col=4, min_row=1, max_col=4, max_row=len(barData))
#     labels = Reference(barSheet, min_col=1, min_row=2, max_row=len(barData))

#     barchart.add_data(data, titles_from_data=True)
#     barchart.set_categories(labels)
#     barchart.legend = None
    
#     barSheet.add_chart(barchart, "H8")
    
#     wb.save(saveas)
    # stop = time.time()
    # print(stop-start)
    


class MainWindow():
    def __init__(self, geometry='900x600') -> None:
        # self.mutex = win32event.CreateMutex(None, 1, "YourAppMutexName")
        # if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        #     # If another instance is already running, exit
        #     print("Another instance of the application is already running.")
        #     return
        
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

    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS2
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)
    

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
        try:
            image = ImageTk.PhotoImage(Image.open( os.path.dirname(sys.executable) +("/assets/logo.jpg")))
        except:
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
        # try:
        #     icon = ImageTk.PhotoImage(Image.open( os.path.dirname(sys.executable) +("/assets/icon.png")))
        # except:
        #     icon = ImageTk.PhotoImage(Image.open("icon.png"))

        self.window.geometry(self.geometry)
        self.window.title('HybridQC')
        self.window.resizable(width=False, height=False)
        self.window.config(bg=self.bg)
        self.center_window(self.window)
        # self.window.iconphoto(True,icon)

        self.toplabel = Label(self.window, text='HybridQC',
                              font=('Georgia', 32, 'bold'),
                              fg='black', bg=self.bg)
        self.toplabel.pack(pady=(30, 40))
        
        self.fileEntry = Entry(self.window,
                               font=('Arial', 16),
                               fg='white',
                               bg=self.bg, highlightthickness=2,
                               width=40, highlightbackground = "black", highlightcolor= "black")
        self.fileEntry.pack()
        
        self.info = Label(self.window,
                         text='Select xlsx file',
                         font=('Arial', 15, 'bold'),
                         fg='black', bg=self.bg)
        self.info.pack(pady=(0,5))
        
        self.selectFileButton = Button(font=('Arial', 12, 'bold'),
                        # fg=self.bg,
                        bg='black', highlightcolor='black',
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
        self.pb.pack(pady=(3,3))

        # progressbar.start()
        watermark = Label(text="© AJIBADE Y.A & Ongom P.O, IITA, Nigeria. 2024",
                          font=("Times New Roman",14,'bold'), bg='white', fg='black')
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
        # self.window.update_idletasks()

        
    
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
            self.log.config(text='Working...', fg='black',)
            
            perc_poly = self.perc_polymorphic_threshold.get()
            perc_missing = self.perc_missing_threshold.get()
            perc_hybrid = self.perc_hybridity_threshold.get()
            
            perc_poly = self.validateInput(perc_poly, 20)
            perc_missing = self.validateInput(perc_missing, 20)
            perc_hybrid = self.validateInput(perc_hybrid, 50)

            savepath = self.saveas()
            if savepath is not None:
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
                self.citationWindow()
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
            
            self.perc_hybridity_threshold.config(state='readonly', fg='white')
            self.perc_missing_threshold.config(state='readonly', fg='white')
            self.perc_polymorphic_threshold.config(state='readonly', fg='white')
            
        else:
            self.perc_hybridity_threshold.config(state='normal', fg='black')
            self.perc_missing_threshold.config(state='normal', fg='black')
            self.perc_polymorphic_threshold.config(state='normal', fg='black')
    
    
    def create_entry_row(self, window):
        bg = '#FFFFCC'
        '#fffc9d'
        row_frame = Frame(window, bg=bg, bd=2, relief=SOLID,
                          width=400, highlightbackground="#fffc9d",
                          height=353)  # Create a frame for the row
        #ffffcc
        Label(row_frame, font=('Arial', 18, 'bold'),
              text="Input Thresholds",fg='black',
              bg=bg).grid(row=0, column=0, columnspan=3, pady=(10, 20))

        Label(row_frame, font=('Arial', 16,),
              text="Min % polymorphic:",fg='black', bg=bg).grid(row=1, column=0, padx=(29, 0))
        self.perc_polymorphic_threshold = Entry(row_frame, font=('Arial', 15),
                                                bg=self.bg, highlightthickness=1,
                                                highlightbackground = "black", highlightcolor= "black",
                                                fg='black', width=5, insertbackground='black',)
        self.perc_polymorphic_threshold.grid(row=1, column=1, padx=20, pady=10)

        Label(row_frame, font=('Arial', 16,),
              text="Max % missing:", fg='black',bg=bg).grid(row=2, column=0)
        self.perc_missing_threshold = Entry(row_frame, font=('Arial', 15), 
                                            bg=self.bg, highlightthickness=1,
                                                highlightbackground = "black", 
                                                highlightcolor= "black", insertbackground='black',
                                            fg='black', width=5)
        self.perc_missing_threshold.grid(row=2, column=1, padx=20, pady=10)

        Label(row_frame, font=('Arial', 16,),
              text="Min % hybridity:",fg='black', bg=bg).grid(row=3, column=0)
        self.perc_hybridity_threshold = Entry(row_frame, font=('Arial', 15), 
                                              bg=self.bg, highlightthickness=1,
                                                highlightbackground = "black",
                                                highlightcolor= "black", insertbackground='black',
                                                fg='black', width=5)
        self.perc_hybridity_threshold.grid(row=3, column=1, padx=20, pady=10)

        self.var = BooleanVar()
        self.defaultThresholdBtn = Checkbutton(row_frame, text="Use default threshold",
                                    variable=self.var, bg=bg, fg='black',
                                    command=self.defaultThreshold)
        self.defaultThresholdBtn.grid(row=4, column=0, pady=10)
        
    
        self.runButton = Button(row_frame, font=('Arial', 15, 'bold'),
                width=18, height= 1,
                   text='Run', 
                   highlightbackground='black',
                   #fg=self.bg,
                #    pady=(15, 5),
                   relief=SOLID,
                   cursor='hand1',
                   command=self.dotheJob)
        self.runButton.grid(row=5, column=0, columnspan=3, pady=(10, 20))
        
        row_frame.pack(pady=(40, 0))  # Pack the entire row frame
        
        
        def citationWindow(self):
            window = self.window
            popup = Toplevel(window)
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
            
            reference_entry.insert(0, "Patrick O., Yakub A.")
            reference_entry.config(state='readonly')
            reference_entry.pack(padx=15)
            

            def close():
                get_referrence = reference_entry.get()
                popup.destroy()
                
            closeBtn = Button(popup, text='CLOSE', command=close)
            closeBtn.pack(pady=(20,0))
            
            
a = MainWindow()
# a.showWindow()
# a.update_button_position(None)
# a.fileEntry.bind('<Configure>', a.update_button_position)

