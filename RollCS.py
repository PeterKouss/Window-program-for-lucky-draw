import pandas as pd
import math
import random
import os
from tkinter import *
from tkinter import messagebox
from tkinter import StringVar, IntVar
import time
import threading

class Participant:
    def __init__(self):
        self.path = "./participant.xlsx"
        self.PPlist = self.readPPExcel()
    
    def readPPExcel(self):
        df = pd.read_excel(self.path)
        df_list = df.values.tolist()
        name = []
        for s_list in df_list:
            name.append(s_list[1])
        return name

    def randomPPlist(self):
        return random.shuffle(self.PPlist)


class Prize:
    def __init__(self):
        self.path = "./prize.xlsx"
        self.PZlist = self.readPZExcel()

    def readPZExcel(self):
        df = pd.read_excel(self.path)
        df_list = df.values.tolist()
        prizename = []
        for s_list in df_list:
            if s_list[1] != None:
                prizename.append(s_list[1])
        return prizename

    def randomPZlist(self):
        return random.shuffle(self.PZlist)


class rollName:
    def __init__(self):
        self.createThread()
        self.beginTread()

    def createThread(self):
        self.t = threading.Thread(target=self.refreshLable)

    def refreshLable(self):
        while(1):
            for _,name in enumerate(DealPP.PPlist):
                eventSig.wait()
                nameVar.set(name)
                root.update()
                time.sleep(0.1)
                
        
    def beginTread(self):
        self.t.start()
      

class Application(Frame):
    def __init__(self, master = None):
        super().__init__(master) 
        self.master=master
        self.pack()
        self.creatWidget()
        self.rollflag=0
        self.prizenumber = 0
        self.allprize = len(DealPZ.PZlist)
        self.winner_list=[]
        self.prize_list=[]
        self.pirze_winner_list=[]

    def creatWidget(self):
        """create widget"""
    
        #create name lable
        nameVar.set("人名滚动")
        self.label01 = Label(self, textvariable = nameVar, width = 20,
                             height = 2, bg = "white", fg = "blue")
        self.label01.pack()
        
        #create prze label
        prizeVar.set("当前奖品")
        self.label02 = Label(self, textvariable = prizeVar, width = 20,
                             height = 2, bg = "white", fg = "blue")
        self.label02.pack()

        #create begin and stop
        self.btn01 = Button(self)
        self.btn01["text"] = "开始/停止"
        self.btn01.pack()
        self.btn01["command"] = self.roll_or_not

        #create show winner button
        self.btn02 = Button(self)
        self.btn02["text"] = "显示获奖列表"
        self.btn02.pack()
        self.btn02["command"] = self.show_winner_list

        #create exit
        self.btnQuit = Button(self, text="退出", command = root.destroy)
        self.btnQuit.pack()

        #create text print
        #text
        self.text01=Text(root, width=30, height=30, undo=True, autoseparators=False)
        self.text01.pack()

    def roll_or_not(self):
        if self.prizenumber<len(DealPZ.PZlist):
            if self.rollflag == 0:
                DealPP.randomPPlist()
                self.prizename = DealPZ.PZlist[self.prizenumber]
                prizeVar.set("Prize No."+str(self.prizenumber+1)+": "+self.prizename)
                eventSig.set()
                self.rollflag = 1
            else:
                eventSig.clear()
                winner = nameVar.get()
                messagebox.showinfo("Congratulation!", winner+" 获得了 "+self.prizename)
                self.prize_list.append(self.prizename)
                self.winner_list.append(winner)
                self.prizenumber += 1
                self.rollflag = 0
        else:
            if self.rollflag == 0:
                DealPP.randomPPlist()
                prizeVar.set("自由抽奖奖品"+str(self.prizenumber))
                self.prizename = "自由抽奖奖品"+str(self.prizenumber)
                eventSig.set()
                self.rollflag = 1
            else:
                eventSig.clear()
                winner = nameVar.get()
                messagebox.showinfo("Congratulation!", winner+" 获得了 "+self.prizename)
                self.prize_list.append(self.prizename)
                self.winner_list.append(winner)
                self.prizenumber += 1
                self.rollflag = 0

    def show_winner_list(self):
        self.pirze_winner_list = list(zip(self.prize_list,self.winner_list))
        self.text01.delete(1.0, "end")
        self.text01.insert(INSERT, "各奖项获得者:\n")
        for i in range(0,len(self.prize_list)):
            self.text01.insert(INSERT, ', '.join(self.pirze_winner_list[i])+"\n")


if __name__ == '__main__':
    
    #global
    DealPP = Participant()
    DealPZ = Prize()  
    eventSig = threading.Event()
    eventSig.clear()
    root = Tk()
    nameVar = StringVar()
    prizeVar = StringVar()
    rollControl=rollName()
    app = Application(master = root)

    #random prize list
    DealPZ.randomPZlist() 

    #genrate Window  
    root.geometry("400x1000")
    root.title("UCAS Rolling for CSGO")

    #print intial information
    app.text01.insert(INSERT, "所有参与者:"+', '.join(DealPP.PPlist)+"\n")
    app.text01.insert(INSERT, "所有奖品:"+', '.join(DealPZ.PZlist)+"\n")

    root.mainloop()



