# -*- coding: utf-8 -*-
"""
Created on Tue Mar 22 13:27:45 2022

@author: Kirsch
"""

import pandas as pd
import tkinter  as tk 
from tkinter import filedialog, ttk
from tkinter.scrolledtext import ScrolledText
import xml.etree.ElementTree as ET


#Error handling class
class error :
    def __init__(self):
        self.er = None
        
    #If wrong files are inputted in the app     
    def open_popup(self, win):
        top= tk.Toplevel(win)
        #top.eval('tk::PlaceWindow . center')

        top.geometry("270x50")
        top.title("Wrong files")
        tk.Label(top, text= "Wrong file inputted, please verify and try again"
                 ,).place(x=0,y=25)
       
        #Reset file and directory name
        l2.config(text = "")
        txt1.config(text = "")
        txt2.config(text = "")
        check_zone.delete("1.0","end")
        inputtxt.delete(0,"end")
    
    #If the billing product file is wrong
    def checkbox(self, win): 
        top= tk.Toplevel(win)
        top.geometry("270x50")
        top.title("Wrong files")
        tk.Label(top, text= "Wrong file inputted, please verify and try again"
                 ,).place(x=0,y=25)
       
        #Reset file and directory name
        l2.config(text = "")
        txt1.config(text = "")
        txt2.config(text = "")
        check_zone.delete("1.0","end")
        inputtxt.delete(0,"end")        

#Class responsible for reading input files 
class input_file():
    def __init__(self, sheet):
        self.df1 = None
        self.sheet = sheet
    
    #Function for XML files
    def upload_file_1(self,txt): 
        try : 
            #Read XML file
            f_types = [('XML files',"*.xml"),('All',"*.*")]
            file = filedialog.askopenfilename(filetypes=f_types)
            txt.config(text = "Ihr Dokument ist:\n" + file, font=my_font2)
            
            #Parse the XML data into a pandas dataframe
            
            xml_file = ET.parse(file)
            self.df1 = pd.DataFrame(columns=['userid','productid', 'item', 'vp_excl', 'datestart', 'dateend'])
    
            root = xml_file.getroot()
            
            for i in range(len(root[0])) : 
                for j in range(len(root[0][i])):
                    row = [None]*len(self.df1.columns)
                    row[0] = root[0][i].attrib["userid"]    
                    row[1] = root[0][i][j].attrib["productid"]
                    row[2] = root[0][i][j].attrib["item"]
                    row[3] = root[0][i][j].attrib["vp_excl"]
                    row[4] = root[0][i][j].attrib["datestart"]
                    row[5] = root[0][i][j].attrib["dateend"]
                    self.df1.loc[len(self.df1)] = row
        except FileNotFoundError : 
            txt.config(text = "No file uploaded \n try again")
                
                
    #Function to read the billing product excel file into a dataframe            
    def upload_file_2(self,txt):
        try : 
            f_types = [('Excel files',"*.xlsx"),('All',"*.*")]
            file = filedialog.askopenfilename(filetypes=f_types)
            txt.config(text = "Your file is :\n" + file, font=my_font2)
            
            excel = pd.ExcelFile(file) # create DataFrame
            self.df1 = pd.read_excel(excel, self.sheet)
            
        except FileNotFoundError : 
            txt.config(text = "No file uploaded \n try again")
        
#Class for output directory
class directory():
    def __init__(self):
        self.directory = None 

    def chose_target(self, l2):
        #Read the directory and display it 
        directory = filedialog.askdirectory()
        self.directory = directory
        l2.config(text = "Die Datei wird im folgenden Ordner gespeichert:\n" + directory, font=my_font2)

#Class for most of the data handling process
class update():  
    #Recuperate all prvious dataframes
    def make_change(self,df1,df2,direc, row_var,row_name,row_ids, l2, txt1,
                    txt2, check_zone, inputtxt, df3, txt1b):
        #if df1 != None & df2 != None &  : 
        try : 
            
            #Merge both billing details files
            df1 = df1.append(df3)
            df1.reset_index(inplace = True)
            
            #Get name of the output file (written by the user)
            inp = inputtxt.get()
            
            #Associate product ID number to product ID name
            dico_prod = dict(zip(df2.fldBillingProductID,df2.fldName))
            dico_prod = {str(k):v for k,v in dico_prod.items()}
            
            #Associate product ID number to product ID name
            dico_order = dict(zip(df2.fldName,df2.fldBillGroup))
            dico_order = {str(k):v for k,v in dico_order.items()}
            
            
            #Associate link checkbox status, name and values in two dictionary
            row_name_val = [None]*len(row_ids)
            row_ids_val = [None]*len(row_ids)
            row_var_val =[None]*len(row_ids)
            
            for i in range(len(row_ids)):    
                row_name_val[i] = row_name[i].get()
                row_ids_val[i] = row_ids[i].cget("text")
                row_var_val[i] = row_var[i].get()
                
            dico_pid = dict(zip(row_ids_val,row_name_val))
            dico_var = dict(zip(row_ids_val,row_var_val))
                
            #Filling output
            erp = pd.DataFrame(columns=['Zeilenart', 'Document Type', 'Ext. Belegnr.','Posting Date',
                                        'Document Date Customer No','Revenue Type Code', 'Erstzeile/Folgezeile',
                                        'Buchungstext', 'Line Identifier', 'Line Quantity','Line unit price',
                                        'Attribute 1','Attribute 2','Attribute 3','Attribute 4','Attribute 5',
                                        'Attribute 6','Attribute 7','Attribute 8','Attribute 9 Layout Identifier',
                                        'External Reference Type', 'External Reference No','Valid from','Valid to',
                                       'External Object Identifier', 'Prepayment Quantity', 'Preisfaktor', 'Statistikcode'])
            
            #Aplly simple data transfers from billing details
            erp["External Reference No"] = df1["userid"]
            erp['Valid from'] = df1["datestart"]
            erp['Valid to'] = df1["dateend"]
            erp["Line Quantity"] = df1["item"]
            erp['Line unit price'] = df1["vp_excl"]
            erp["Zeilenart"] = 4
            erp["Document Type"] = 2
            erp["Revenue Type Code"] = "GFA-L"
            erp["Preisfaktor"] = 1
            erp['Buchungstext'] = df1['productid']
            erp['Erstzeile/Folgezeile'] = 2
            
            
            #Convert datetime formats to the right one
            erp['Valid from'] = erp['Valid from'].str.split("T",1).str[0]
            erp['Valid from'] = pd.to_datetime(erp['Valid from'])
            erp['Valid from'] = erp['Valid from'].dt.strftime('%d.%m.%Y')
            
            erp['Valid to'] = erp['Valid to'].str.split("T",1).str[0]
            erp['Valid to'] = pd.to_datetime(erp['Valid to'])
            erp['Valid to'] = erp['Valid to'].dt.strftime('%d.%m.%Y')
        
        
        
            #Replace product id code to its name
            erp.replace({"Buchungstext": dico_prod}, inplace = True)
            
            
            erp['Line Identifier'] = erp["Buchungstext"]
            erp['value'] = erp["Buchungstext"]
            erp["order"] = erp["Buchungstext"]
            

            erp.replace({'order': dico_order}, inplace = True)

            orders = {11:1,1:2,2:3,4:4,8:5,9:6,10:7,7:8,5:9,6:10}
            erp.replace({"order": orders}, inplace = True)
            
            
            erp.loc[erp["Buchungstext"] == "QL_SMARTACCESS", "order"] = 0
            
            
            #Add checkbox status and value to all lines
            erp.replace({'Line Identifier': dico_pid}, inplace = True)
            erp.replace({'value': dico_var}, inplace = True)
            erp.replace({'order': dico_order}, inplace = True)
            
            #Add "GFA-L'
            erp["Buchungstext"] = "GFA-L " + erp["Buchungstext"].astype(str)
            
            
            #Sort the output file by User ID, type of operation and operation month
            erp.sort_values(["External Reference No", "order", 
                            "Valid from"],
               axis = 0, ascending = True,
               inplace = True)          

            #Drop checked lines with unit price equal to zero
            erp = erp[(erp["Line unit price"].astype(float) != 0) | (erp["value"] != 1)]
            erp.drop(["value", "order"], axis = 1, inplace = True) 


            #Do the Ertzeile/Fortzeile split 
            ids = list(df1.userid.unique())
            for i in ids:
                e = erp["External Reference No"].ne(i).idxmin()
                erp["Erstzeile/Folgezeile"][e] = 1
                
           
            #Save the file to excel 
            erp.to_excel(direc + "/" + inp + ".xlsx", index=False)
            
            
            final_msg = tk.Label(text = "Die Datei wurde erfolgreich in folgendem Ordner erstellt:  \n" + direc)
            final_msg.place(relx = 0.5, rely = 0.9, anchor = "center")
            
            #Reset file and directory name
            l2.config(text = "")
            txt1.config(text = "")
            txt2.config(text = "")
            txt1b.config(text = "")
            check_zone.delete("1.0","end")
            inputtxt.delete(0,"end")
            
        except AttributeError: 
            err.open_popup(my_w)
            
                        
#Manage checkboxes for billing products
class checkboxes(): 
    def __init__(self):
        self.ids = [] 
        self.var = [] 
        self.rows = []


    def select_all(self, ids):
        for j in ids:
            j.select()
            
    def deselect_all(self, ids): 
         for j in ids:
            j.deselect()
        

    def checkboxes(self, text, df): 
        try : 
            #Get all the Product names
            names = df["fldName"]
            text.insert("end", "Positionen, die nicht importiert werden sollen markieren \n\n")
            text.insert('end', '\t')
            
            #Button to select all checkboxes
            sel_all = tk.ttk.Button(text, text="select all",
                                    command = lambda:self.select_all(self.ids))
            text.window_create('end', window=sel_all)
            text.insert('end', '\t \t ')
    
            #Button to deslct all 
            desel_all = tk.ttk.Button(text, text="deselect all",
                    command = lambda:self.deselect_all(self.ids))
            text.window_create('end', window=desel_all)
            
            text.insert('end', '\n')
            
            #Create a checkbox per product 
            for i in range(len(names)):
                variable = tk.IntVar()
                self.var.append(variable)
                cb = tk.Checkbutton(text, text=str(names[i]),
                                        var = variable)
                
                
                
                self.ids.append(cb)
                text.window_create('end', window=cb)
                v = tk.StringVar(text, value='8900.040')
                self.rows.append(v)
                value1 = tk.Entry(text, textvariable = v) 
                text.window_create('end', window=value1)
                text.insert('end', '\n')
            
        except TypeError : 
            error.checkbox(self, my_w)
        
#Instantiate the classes
a = input_file("Tabelle1")   
m = input_file("Tabelle1") 
b = input_file("Aktuelle")
d = directory()
c = update()
check = checkboxes()
err = error()

#Instantiate the window
my_w = tk.Tk()
my_w.geometry("1000x9000")  # Size of the window 
my_w.title('GFA Luthern')


my_w.iconbitmap("gfa_luth.ico")



#Create header
my_font1=('calibri', 30, 'bold')
my_font2= ("calibri", 10, 'italic') 
l1 = tk.ttk.Label(my_w,text='Willkommen!',
    width=12,font=my_font1)  
l1.place(relx=0.5, rely = 0.05, anchor='center')


#Upload first billing detail file button 
b1 = tk.ttk.Button(my_w, text='Billing Details Monat 1 importieren', 
   width=35,command = lambda:a.upload_file_1(txt = txt1))
b1.place(relx=0.25, rely = 0.2, anchor='center') 

#Text with first file uploaded
txt1 = tk.Label()
txt1.place(relx=0.25, rely=0.25, anchor='center')

#Upload second billing detail file button 
b1b = tk.ttk.Button(my_w, text='Billing Details Monat 2 importieren', 
   width=35,command = lambda:m.upload_file_1(txt = txt1b))
b1b.place(relx=0.25, rely = 0.3, anchor='center') 

#Text with second file uploaded
txt1b = tk.Label()
txt1b.place(relx=0.25, rely=0.35, anchor='center')


#Upload billing product file
b2 = tk.ttk.Button(my_w, text='Billing Products Importieren', 
   width=30,command = lambda:b.upload_file_2(txt = txt2))
b2.place(relx=0.25, rely = 0.4, anchor='center')

#Text with billing product file
txt2 = tk.Label()
txt2.place(relx=0.25, rely=0.45, anchor='center')




#Variables to track if checked or not
var1 = tk.IntVar()
var2 = tk.IntVar()
var = [var1,var2]

#Create the checkbox zone for billing product selection
check_zone = ScrolledText(my_w, width=40, height=35)
check_zone.place(relx=0.75, rely = 0.5, anchor='center')


#Button to make the billing product appear
row_button = tk.ttk.Button(my_w, text = 'ID\'s Prüfen', width = 20, command = lambda:check.checkboxes(check_zone,b.df1))
row_button.place(relx=0.75, rely = 0.15, anchor='center')


l2 = tk.Label()
l2.place(relx=0.25, rely = 0.6, anchor='center')

#Button for target directory
b2 = tk.ttk.Button(my_w, text='Zielspeicherort wählen', 
   width=20,command = lambda:d.chose_target(l2 = l2))   
b2.place(relx=0.25, rely = 0.55, anchor='center')


l3 = tk.Label(my_w, text = "Dateiname eingeben : ", font = 'calibri 10 bold')
l3.place(relx=0.25, rely = 0.67, anchor='center')
#Enter name of the output file
inputtxt = tk.ttk.Entry(my_w)
inputtxt.place(relx=0.25, rely = 0.7, anchor='center')
               

boldStyle = tk.ttk.Style()
boldStyle.configure("Bold.TButton", font = ('Sans','10','bold'))


#Proceed to the data manipulation
b3 = tk.ttk.Button(my_w, text='Rechnungsdaten \n  generierieren', 
                   style = "Bold.TButton",
   width=30, command = lambda:c.make_change(a.df1, b.df1, d.directory,
                                           check.var, check.rows, check.ids,
                                           l2, txt1, txt2, check_zone,
                                           inputtxt, m.df1, txt1b))
b3.place(relx = 0.5, rely = 0.85, anchor = "center")


if __name__ == '__main__':
    my_w.mainloop()


