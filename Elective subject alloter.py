import pandas as pd
import tkinter.messagebox
import customtkinter
from tkinter import filedialog

#setting default appearence as 'SYSTEM' and buttons colour to 'BLUE'
customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("blue")

global title
title = "Open Elective Subject Allotement System.exe"
class EXIT(customtkinter.CTk):
        #exiting window
        def __init__(self, op_path):
                self._op_path=op_path
                super().__init__()
        
                self.title(title)
                self.geometry(f'{800}x{300}')
                self.resizable(0,0)

                self.grid_columnconfigure((0,2), weight=0)
                self.grid_columnconfigure((1), weight=1)
                self.grid_rowconfigure((0, 1, 2, 3, 4), weight=0)
        
                self.instruct1=customtkinter.CTkLabel(self, text="Process Done", font=customtkinter.CTkFont(size=20, weight="bold"))
                self.instruct1.grid(row=1, column=1, sticky='nsew', pady=25)
        
                self.instruct2=customtkinter.CTkLabel(self, text="File Saved as 'Allocation.xlsx' at", font=customtkinter.CTkFont(size=15))
                self.instruct2.grid(row=2, column=1, sticky='nsew', pady=10)
                
                self.file1_entry = customtkinter.CTkEntry(self)
                self.file1_entry.grid(row=3, column=1, sticky='nsew', pady=5)
                self.file1_entry.insert(0, op_path)
                
                self.end_button=customtkinter.CTkButton(self, text="OPEN", command=self.openfile)
                self.end_button.grid(row=4, column=0, sticky='e', pady=20)

                self.end_button=customtkinter.CTkButton(self, text="EXIT", command=self.destroy)
                self.end_button.grid(row=4, column=2, sticky='w', pady=20)
        
        def openfile(self):
                import os
                os.startfile(self._op_path)
        
class Front(customtkinter.CTk):
        #main window
        def __init__(self):
                super().__init__()
                
                self.title(title)
                self.geometry(f'{800}x{300}')
                self.resizable(0,0)

                self.grid_columnconfigure((0,2), weight=0)
                self.grid_columnconfigure((1), weight=1)
                self.grid_rowconfigure((0, 1, 2, 3, 4, 5, 6), weight=0)
                
                self.sidebar_frame = customtkinter.CTkFrame(self, corner_radius=10)
                self.sidebar_frame.grid(row=6, column=0, sticky='w')
                self.sidebar_frame.grid_rowconfigure((0, 1), weight=1)

                self.instruct=customtkinter.CTkLabel(self, text="ELECTIVE ALLOCATION SYSTEM", font=customtkinter.CTkFont(size=15, weight="bold"))
                self.instruct.grid(row=0, column=1, sticky='nsew', pady=20)
                
                self.file1_label = customtkinter.CTkLabel(self, text="Select Student Choice :")
                self.file2_label = customtkinter.CTkLabel(self, text="Select Subject with Branch :")
                self.address_label = customtkinter.CTkLabel(self, text="Select Address to Store output :")

                self.file1_entry = customtkinter.CTkEntry(self)
                self.file2_entry = customtkinter.CTkEntry(self)
                self.address_entry = customtkinter.CTkEntry(self)

                self.file1_browse_button = customtkinter.CTkButton(self, text="Browse", command=self.browse_file1)
                self.file2_browse_button = customtkinter.CTkButton(self, text="Browse", command=self.browse_file2)
                self.folder_browse_button = customtkinter.CTkButton(self, text="Browse", command=self.browse_folder)

                self.submit_button = customtkinter.CTkButton(self, text="Submit", command=self.submit)
                self.instruction_button = customtkinter.CTkButton(self, text="Help", command=self.instruction)
                self.cread_button = customtkinter.CTkButton(self, text="Creadit", command=self.cread)

                self.file1_label.grid(row=1, column=0, padx=2, pady=2, sticky="w")
                self.file1_entry.grid(row=1, column=1, padx=2, pady=2, sticky="we")
                self.file1_browse_button.grid(row=1, column=2, padx=2, pady=2, sticky="e")
                self.file2_label.grid(row=2, column=0, padx=2, pady=2, sticky="w")
                self.file2_entry.grid(row=2, column=1, padx=2, pady=2, sticky="we")
                self.file2_browse_button.grid(row=2, column=2, padx=2, pady=2, sticky="e")
                self.address_label.grid(row=3, column=0, padx=2, pady=2, sticky="w")
                self.address_entry.grid(row=3, column=1, padx=2, pady=2, sticky="we")
                self.folder_browse_button.grid(row=3, column=2, padx=2, pady=2, sticky="e")
                self.submit_button.grid(row=5, column=1, padx=2, pady=2, sticky="n")
                self.instruction_button.grid(row=6, column=1, pady=5, sticky="e")
                self.cread_button.grid(row=6, column=2, pady=5, sticky="e")
                
                self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Appearance Mode")
                self.appearance_mode_label.grid(row=0, padx=5)
                self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["System","Light", "Dark"],command=self.change_appearance_mode_event)
                self.appearance_mode_optionemenu.grid(row=1, padx=5, sticky="n")
                
        def browse_file1(self):
                file_path = filedialog.askopenfilename()
                self.file1_entry.delete(0, tkinter.END)
                self.file1_entry.insert(0, file_path)

        def browse_file2(self):
                file_path = filedialog.askopenfilename()
                self.file2_entry.delete(0, tkinter.END)
                self.file2_entry.insert(0, file_path)

        def browse_folder(self):
                folder_path = filedialog.askdirectory()
                self.address_entry.delete(0, tkinter.END)
                self.address_entry.insert(0, folder_path)
                
        def change_appearance_mode_event(self, new_appearance_mode: str):
                customtkinter.set_appearance_mode(new_appearance_mode)
                
        def cread(self):
                #creadit window
                self = customtkinter.CTk()
                self.title(title)
                self.geometry(f'{700}x{400}')
                self.resizable(0,0)
        
                self.grid_columnconfigure((1), weight=1)
                self.grid_rowconfigure((0, 1, 2, 3, 4, 5, 6, 7, 8 ,9), weight=0)
        
                self.instruct=customtkinter.CTkLabel(self, text="CREDITS", font=customtkinter.CTkFont(size=20, weight="bold"))
                self.instruct.grid(row=0, column=1, sticky='nsew', pady=20)
                
                prt0="Students of ECE TY Batch 2023-24"
                prt1="Tejas Narendra Varute."
                prt2="Shripad Hanmantsing Rajput"
                prt3="Swapnil Balasaheb Shinde"
                prt4="Parth Dhananjay Shinde"
                prt5="Under the Guidence Of"
                prt6="Mr. M. M. Babar"

                self.instruct1=customtkinter.CTkLabel(self, text=prt0, font=customtkinter.CTkFont(size=14, weight="bold"))
                self.instruct1.grid(row=1, column=1, sticky='we', pady=10)
                
                self.instruct1=customtkinter.CTkLabel(self, text=prt1, font=customtkinter.CTkFont(size=12, weight="bold"))
                self.instruct1.grid(row=2, column=1, sticky='we')

                self.instruct2=customtkinter.CTkLabel(self, text=prt2, font=customtkinter.CTkFont(size=12, weight="bold"))
                self.instruct2.grid(row=3, column=1, sticky='we')

                self.instruct3=customtkinter.CTkLabel(self, text=prt3, font=customtkinter.CTkFont(size=12, weight="bold"))
                self.instruct3.grid(row=4, column=1, sticky='we')

                self.instruct4=customtkinter.CTkLabel(self, text=prt4, font=customtkinter.CTkFont(size=12, weight="bold"))
                self.instruct4.grid(row=5, column=1, sticky='we')
        
                self.tmp=customtkinter.CTkLabel(self,text='')
                self.tmp.grid(row=6, column=1, sticky='we', pady=5)
        
                self.instruct5=customtkinter.CTkLabel(self, text=prt5, font=customtkinter.CTkFont(size=14, weight="bold"))
                self.instruct5.grid(row=7, column=1, sticky='we')
        
                self.instruct6=customtkinter.CTkLabel(self, text=prt6, font=customtkinter.CTkFont(size=12, weight="bold"))
                self.instruct6.grid(row=8, column=1, sticky='we')

                self.end_button=customtkinter.CTkButton(self, text="OK", command=self.destroy)
                self.end_button.grid(row=9, column=1, sticky='s', pady=40)
                self.mainloop()

        def instruction(self):
                #instruction window
                self = customtkinter.CTk()
                self.title(title)
                self.geometry(f'{700}x{400}')
                self.resizable(0,0)
                
                self.grid_columnconfigure((0, 1), weight=1)
                self.grid_rowconfigure((0, 1, 2, 3, 4, 5, 6, 7), weight=0)

                self.tmp=customtkinter.CTkLabel(self,text='')
                self.tmp.grid(row=0, column=1, sticky='nsew', pady=10)
        
                self.instruct=customtkinter.CTkLabel(self, text="Help Instructions", font=customtkinter.CTkFont(size=20, weight="bold"))
                self.instruct.grid(row=1, column=1, sticky='ew', pady=30)

                prt1="1. To create student choice data use google form to get information from students."
                prt2="2. To create subject with branch file use refence of give sample file."
                prt3="3. Set the column names as same as given sample file as it is Case sensitive."
                prt4="4. Choose proper file while browsing."
                prt5="5. You can be set number of choices, subjects, branches."

                self.instruct1=customtkinter.CTkLabel(self, text=prt1, font=customtkinter.CTkFont(size=12, weight="bold"))
                self.instruct1.grid(row=2, column=1, sticky='w')

                self.instruct2=customtkinter.CTkLabel(self, text=prt2, font=customtkinter.CTkFont(size=12, weight="bold"))
                self.instruct2.grid(row=3, column=1, sticky='w')

                self.instruct3=customtkinter.CTkLabel(self, text=prt3, font=customtkinter.CTkFont(size=12, weight="bold"))
                self.instruct3.grid(row=4, column=1, sticky='w')

                self.instruct4=customtkinter.CTkLabel(self, text=prt4, font=customtkinter.CTkFont(size=12, weight="bold"))
                self.instruct4.grid(row=5, column=1, sticky='w')
                
                self.instruct4=customtkinter.CTkLabel(self, text=prt5, font=customtkinter.CTkFont(size=12, weight="bold"))
                self.instruct4.grid(row=6, column=1, sticky='w')

                self.end_button=customtkinter.CTkButton(self, text="OK", command=self.destroy)
                self.end_button.grid(row=7, column=1, sticky='s', pady=50)
                self.mainloop()
                
        def submit(self):
                lst_path = self.file1_entry.get()
                sub_path= self.file2_entry.get()
                op_path = self.address_entry.get()
                
                self.destroy()
                allocate=allocation()
                allocate.allocationprocess(lst_path, sub_path, op_path)

class allocation():
    def allocationprocess(self, lst_path, sub_path, op_path):
        #getting Files
        df1 = pd.read_excel(lst_path)
        df2 = pd.read_excel(sub_path)

        #sorting
        temp_df1 = df1.sort_values(by='Timestamp', ascending=False)
        temp_df1 = temp_df1.sort_values(by='First Name', ascending=True)
        temp_df1 = temp_df1.sort_values(by='Last Name', ascending=True)
        student_data = temp_df1.sort_values(by='CGPA', ascending=False)
       

        #List Declarations
        sub,branch,f_name,m_name,l_name,roll,div,subject,st_branch,st_cgpa,prn,seats=[[] for i in range (12)]
        
        #Getting Subject and Branches
        for _, ele in df2.iterrows():
                sub.append(ele['Subject'])
                branch.append(ele['Branch'])
                seats.append(ele['Seats'])

        #Declaring Seats per subject as User given
        subjects_capacity = {sub[i]: seats[i] for i in range(len(sub))}
        subjects_branch ={sub[i] : branch[i] for i in range (len(sub))}

        #Alloting
        for _, student in student_data.iterrows():
                for i in range(len(sub)): 
                        subject_choice = student[f' [Choice {i+1}]']
                        branch=student['Branch']
                        if subjects_capacity[subject_choice] > 0 and subjects_branch[subject_choice] != branch and student['PRN'] not in prn:
                                f_name.append(student['First Name'])
                                m_name.append(student['Middle Name'])
                                l_name.append(student['Last Name'])
                                div.append(student['Div'])
                                st_cgpa.append(student['CGPA'])
                                roll.append(student['Roll No.'])
                                subject.append(subject_choice)
                                st_branch.append(branch)
                                prn.append(student['PRN'])
                                subjects_capacity[subject_choice] -= 1
                                break
        
        Full_name=[f'{f_name[i]} {m_name[i]} {l_name[i]}' for i in range (len(f_name))]
        #Storing to Output File
        op_path=f'{op_path}/Allotement.xlsx'
        wr=[0]*len(sub)
        write = pd.ExcelWriter(path=op_path,engine='xlsxwriter')
        for ele in range(len(sub)):
                temp_name,temp_roll,temp_sub,temp_branch,temp_div,temp_cgpa,temp_prn=[[] for i in range (7)]
                for ele1 in range(len(f_name)):
                        if sub[ele]==subject[ele1]:
                                temp_name.append(Full_name[ele1])
                                temp_roll.append(roll[ele1])
                                temp_div.append(div[ele1])
                                temp_sub.append(subject[ele1])
                                temp_branch.append(st_branch[ele1])
                                temp_cgpa.append(st_cgpa[ele1])
                                temp_prn.append(prn[ele1])
                    
                wr[ele]=(pd.DataFrame({'PRN': temp_prn, 'Division':temp_div, 'Roll No.':temp_roll, 'Branch': temp_branch, 'Name': temp_name, 'CGPA' : temp_cgpa, 'Subject' : temp_sub}))
                wr[ele].to_excel(write,sheet_name=subjects_branch[sub[ele]],index=False)
        write.close()
    
        end = EXIT(op_path)        
        end.mainloop()

front = Front()        
front.mainloop()