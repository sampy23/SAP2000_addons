import os
import sys
import comtypes.client
import time
import shutil
import tkinter as tk
from tkinter import messagebox

class App():
    def __init__(self,master):
        self.master = master

        self.master.title("selection")
        self.font_size = ("Courier", 12)
        self.master.frame_1 = tk.LabelFrame(master)
        self.master.frame_1.grid(row=0,column=0)
        
        title_list = ["Center of building (m):","Length (m):",
                                            "Highest Z Coord (m):"]
        pre_entry_set = ["0,0,0","100","100"]
        nrow = 0
        self.entry_set = []
        for title,entry in zip(title_list,pre_entry_set):
            lbl = tk.Label(root.frame_1, text=title)
            lbl.grid(row=nrow, column=0, sticky='e')
            lbl.config(font=self.font_size) 
            ent = tk.Entry(root.frame_1)
            ent.insert(1,entry)
            ent.grid(row=nrow, column=1)
            self.entry_set.append(ent)
            nrow += 1
        button = tk.Button(root.frame_1,text = "OK",width=8,relief = 'raised',command = self.assign)
        button.grid(row = nrow,column=0,columnspan = 2,padx=10,pady=10)

    def assign(self):
        """Function to create a frame to hold buttons for the operation"""
        self.entry_set = [x.get() for x in self.entry_set]
        self.master.frame_1.destroy()
        # recreating frame_1
        self.master.frame_1 = tk.LabelFrame(root)
        self.master.frame_1.grid(row=0,column=0)
        self.button_1 = tk.Button(self.master.frame_1,text = "X/Y Sn cut",width=16,
                    bg = 'pale green',activebackground = 'orange red',relief = 'raised',command = self.y_sncut)
        self.button_2 = tk.Button(self.master.frame_1,text = "Inclined Sn cut",width=16,
                    bg = 'pale green',activebackground = 'orange red',relief = 'raised',command = self.inclined_sncut)
        self.button_3 = tk.Button(self.master.frame_1,text = "Mirror selection",width=16,
                    bg = 'pale green',activebackground = 'orange red',relief = 'raised',command = self.mirror)
        self.button_4 = tk.Button(self.master.frame_1,text = "Select similar",width=16,
                    bg = 'pale green',activebackground = 'orange red',relief = 'raised',command = self.select_similar)

        self.button_1.grid(row = 0,column=0,padx=10,pady=10)
        self.button_2.grid(row = 0,column=1,padx=10,pady=10)
        self.button_3.grid(row = 1,column=0,padx=10,pady=10)
        self.button_4.grid(row = 1,column=1,padx=10,pady=10)
        self.button_1.config(font=self.font_size)
        self.button_2.config(font=self.font_size)
        self.button_3.config(font=self.font_size)
        self.button_4.config(font=self.font_size)


    def attach_to_instance(self):
        """Function to attach to instance of SAP2000"""
        try:
            #get the active ETABS object
            self.myETABSObject = comtypes.client.GetActiveObject("CSI.SAP2000.API.SapObject") 
        except (OSError, comtypes.COMError):
            self.no_model()

        try:
            self.SapModel = self.myETABSObject.SapModel
        except (OSError, comtypes.COMError):
            self.no_model()

        self.curr_unit = self.SapModel.GetPresentUnits() #kn_m_C
        self.SapModel.SetPresentUnits(6) #kn_m_C
        self.model_path = self.SapModel.GetModelFilename()
        base_name = os.path.basename(self.model_path)[:-4]
        self.backup(self.model_path) # backup function
        return base_name

    def backup(self,model_path):
        """Creates backup for the active file in root directory of SAP2000 file"""
        if 4 in set(self.SapModel.Analyze.GetCaseStatus()[2]): # 4 indicates a run case
            #atleast one case has run so no need to save the file and lose analysis data
            pass
        else:
            self.SapModel.File.Save(model_path) #save file before taking backup, but this deletes program results
        os.chdir(os.path.dirname(model_path))
        # model backup
        try:
            os.mkdir(".//_backup")
        except FileExistsError:
            pass
        file_dir = os.path.dirname(model_path)
        file_name_ext = os.path.basename(model_path)
        file_name,ext = os.path.splitext(file_name_ext)
        time_stamp = time.strftime("%H%M%S")
        new_file_name = file_name+ "_" + time_stamp + ext
        new_model_path = os.path.join("_backup",new_file_name)
        self.new_model_path = os.path.join(file_dir,new_model_path)
        shutil.copy2(model_path,new_model_path) 


    def point_label(self,label):
        joints = self.SapModel.FrameObj.GetPoints(label)
        point_1 = self.SapModel.PointObj.GetCoordCartesian(joints[0])[:-1]
        point_2 = self.SapModel.PointObj.GetCoordCartesian(joints[1])[:-1]
        return point_1,point_2
        
    def select_bound(self,point_1,point_2,x = True,y=True,orgin = (0,0,0)):
        """x= True means x of selected member quadrant and similarly y,in this way we can define the quadrants"""
        zmin = min(point_1[2],point_2[2])
        zmax = max(point_1[2],point_2[2])
        
        kwargs_1 = {"XMin" : min(point_1[0],point_2[0]),\
                    "XMax" : max(point_1[0],point_2[0]),\
                    "YMin" : min(point_1[1],point_2[1]),\
                    "YMax" : max(point_1[1],point_2[1]),\
                    "ZMin" : zmin,\
                    "ZMax" :zmax }
        kwargs_2 = {"XMin" : min(point_1[0],point_2[0]),\
                    "XMax" : max(point_1[0],point_2[0]),\
                    "YMin" : -max(point_1[1],point_2[1])+ 2 * orgin[1],\
                    "YMax" : -min(point_1[1],point_2[1]) + 2 * orgin[1],\
                    "ZMin" : zmin,\
                    "ZMax" : zmax}
        kwargs_3 = {"XMin" : -max(point_1[0],point_2[0])+ 2 * orgin[0],\
                    "XMax" : -min(point_1[0],point_2[0])+ 2 * orgin[0],\
                    "YMin" : min(point_1[1],point_2[1]),\
                    "YMax" : max(point_1[1],point_2[1]),\
                    "ZMin" : zmin,\
                    "ZMax" : zmax}
        kwargs_4 = {"XMin" : -max(point_1[0],point_2[0])+ 2 * orgin[0],\
                    "XMax" : -min(point_1[0],point_2[0])+ 2 * orgin[0],\
                    "YMin" : -max(point_1[1],point_2[1])+ 2 * orgin[1],\
                    "YMax" : -min(point_1[1],point_2[1])+ 2 * orgin[1],\
                    "ZMin" : zmin,\
                    "ZMax" : zmax}
        
        if x == True and y == True:
            self.SapModel.SelectObj.CoordinateRange(**kwargs_1,IncludeIntersections = False)
        elif x == True and y == False:
            self.SapModel.SelectObj.CoordinateRange(**kwargs_2,IncludeIntersections = False)
        elif x == False and y == True:
            self.SapModel.SelectObj.CoordinateRange(**kwargs_3,IncludeIntersections = False)
        elif x == False and y == False:
            self.SapModel.SelectObj.CoordinateRange(**kwargs_4,IncludeIntersections = False)
    @staticmethod    
    def collect_framelabels(selected):
        frame_labels = []
        object_types = selected[1] 
        selected_labels = selected[2] 
        for object_type,label in zip(object_types,selected_labels):
            if object_type == 2:
                frame_labels.append(label)
        return frame_labels

    def filter_unwanted(self,target_point1,target_point2,selected_frame_labels,orgin):
        matching = []
        target_point1_sum = abs(target_point1[0] - orgin[0]) + abs(target_point1[1] - orgin[1]) + abs(target_point1[2] - orgin[2])
        target_point2_sum = abs(target_point2[0] - orgin[0]) + abs(target_point2[1] - orgin[1]) + abs(target_point2[2] - orgin[2])
        target_sum = round(target_point1_sum + target_point2_sum,4)

        for lab in selected_frame_labels:
            pon_1,pon_2 = self.point_label(lab)
            pon_1_sum = abs(pon_1[0] - orgin[0]) + abs(pon_1[1] - orgin[1]) + abs(pon_1[2] - orgin[2])
            pon_2_sum = abs(pon_2[0] - orgin[0]) + abs(pon_2[1] - orgin[1]) + abs(pon_2[2] - orgin[2])
            pon_sum = round(pon_1_sum + pon_2_sum,4) 
            if (target_sum == pon_sum):
                matching.append(lab)
        return matching

    @staticmethod
    def slope(point_1,point_2):
        x1 = point_1[0]
        y1 = point_1[1]
        x2 = point_2[0]
        y2 = point_2[1]
        try:
            slp = (y2-y1)/(x2-x1)
        except ZeroDivisionError:
            slp = float('inf')
        return slp

    def filter_unwanted_slope(self,target_point_1,target_point_2,selected_frame_labels):
        matching = []
        target_slope = round(self.slope(target_point_1,target_point_2),4)
        for lab in selected_frame_labels:
            point_1,point_2 = self.point_label(lab)
            slpe = round(self.slope(point_1,point_2),4)
            point_1 = [round(x,4) for x in point_1]
            point_2 = [round(x,4) for x in point_2]
            if target_slope == slpe:
                matching.append(lab)
            if point_1[:-1] == point_2[:-1]: # to include verticals
                matching.append(lab)
        return matching
        
    def y_sncut(self):
        file_name = self.attach_to_instance()
        self.master.title(file_name)
        # section cut, in X/Y plane 
        object_types = self.SapModel.SelectObj.GetSelected()[1] 
        selected_labels = self.SapModel.SelectObj.GetSelected()[2]
        length = float(self.entry_set[1]) # an arbitrary number greater than length of building
        max_z = float(self.entry_set[2])
        for object_type,label in zip(object_types,selected_labels):
            if object_type == 2: # to exclude points
                joints = self.SapModel.FrameObj.GetPoints(label)
                point_1 = self.SapModel.PointObj.GetCoordCartesian(joints[0])[:-1]
                point_2 = self.SapModel.PointObj.GetCoordCartesian(joints[1])[:-1]
                if round(point_1[0],4) == round(point_2[0],4): #for y section cut
                    self.SapModel.SelectObj.CoordinateRange(min(point_1[0],point_2[0]),max(point_1[0],point_2[0]),\
                                                min(point_1[1],point_2[1])-length,max(point_1[1],point_2[1])+length,\
                                                -max_z,max_z,IncludeIntersections = False)
                if round(point_1[1],4) == round(point_2[1],4): #for x section cut
                    self.SapModel.SelectObj.CoordinateRange(min(point_1[0],point_2[0])-length,max(point_1[0],\
                                                    point_2[0])+length,min(point_1[1],point_2[1]),max(point_1[1],\
                                                    point_2[1]),-max_z,max_z,IncludeIntersections = False)
        self.SapModel.View.RefreshView()
        
    def inclined_sncut(self):
        file_name = self.attach_to_instance()
        self.master.title(file_name)
        # SECTION CUT AT AN ANGLE
        # WARNING !! UNNEEDED VERTICAL WILL BE CHOSEN UNFORTUNATELY ITS A SHORTCOMING OF API SELECTION FUNCTION
        target_labels = []
        max_z = float(self.entry_set[2])
        joints = self.SapModel.SelectObj.GetSelected()[2] #select two bounding points
        point_1 = self.SapModel.PointObj.GetCoordCartesian(joints[0])[:-1]
        point_2 = self.SapModel.PointObj.GetCoordCartesian(joints[1])[:-1]
        self.SapModel.SelectObj.CoordinateRange(min(point_1[0],point_2[0]),max(point_1[0],point_2[0]),\
                                        min(point_1[1],point_2[1]),max(point_1[1],point_2[1])\
                                        ,-max_z,max_z,IncludeIntersections = False)
        selected_frame_labels = self.collect_framelabels(self.SapModel.SelectObj.GetSelected())
        target_labels.extend(self.filter_unwanted_slope(point_1,point_2,selected_frame_labels))
        self.SapModel.SelectObj.ClearSelection()
        for ele in target_labels:
            self.SapModel.FrameObj.SetSelected(ele,True)
        self.SapModel.View.RefreshView()
        
    def select_similar(self):
        file_name = self.attach_to_instance()
        self.master.title(file_name)
        # select frames of similar section
        frame_info = self.SapModel.SelectObj.GetSelected()
        selected_labels = frame_info[2]
        if len(selected_labels) > 1:
            print("More than one elements selected")
            sys.exit()
        target_section = self.SapModel.FrameObj.GetSection(selected_labels[0])[0]
        for label in self.SapModel.FrameObj.GetNameList()[1]:
            if target_section == self.SapModel.FrameObj.GetSection(label)[0]:
                self.SapModel.FrameObj.SetSelected(label,True)
        self.SapModel.View.RefreshView()
        print("selected frame is",target_section)

    def mirror(self):
        file_name = self.attach_to_instance()
        self.master.title(file_name)
        # select elements in mirror location        
        target_labels = []
        center_of_model = [float(x) for x in self.entry_set[0].split(",")]
        if len(center_of_model) != 3: 
            center_of_model[2:] = [0] # assuming we missed or put more than 3 coords
        # center_of_model = (6,11,8.6)
        init_frame_labels = self.collect_framelabels(self.SapModel.SelectObj.GetSelected())
        for label in init_frame_labels:
            point_1,point_2 = self.point_label(label)
            self.select_bound(point_1,point_2,True,True,center_of_model)
            self.select_bound(point_1,point_2,True,False,center_of_model)
            self.select_bound(point_1,point_2,False,False,center_of_model)
            self.select_bound(point_1,point_2,False,True,center_of_model)
            selected_frame_labels = self.collect_framelabels(self.SapModel.SelectObj.GetSelected())
            target_labels.extend(self.filter_unwanted(point_1,point_2,selected_frame_labels,center_of_model))
        self.SapModel.SelectObj.ClearSelection()
        for ele in target_labels:
            self.SapModel.FrameObj.SetSelected(ele,True)
        self.SapModel.View.RefreshView()

    def no_model(self):
        self.master.withdraw()
        messagebox.showwarning(title = "Active model not found",
                        message = "Close all SAP2000 instances if any open and reopen the target file first.")
        messagebox.showinfo(title = "Help",message = "For trouble shooting contact me through sbz5677@gmail.com ")
        self.master.destroy()
        sys.exit()
    
    def return_values(self):
        return self.SapModel,self.curr_unit

if __name__ == '__main__':
    root = tk.Tk()
    inst_1  = App(root)
    root.mainloop()
    # reset to the previous units
    SapModel, curr_unit = inst_1.return_values()
    SapModel.SetPresentUnits(curr_unit)
    # exiting information
    tk.Tk().withdraw()