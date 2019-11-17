from tkinter import ttk,Toplevel,IntVar
from tkinter import Tk as ThemedTk

value : str
counter : int
player_1=[]
player_2=[]

def CustomMsgBox(Message): 
    child = Toplevel(root) 
    child.resizable(False,False)
    windowWidth = root.winfo_reqwidth()
    windowHeight = root.winfo_reqheight()
    positionRight = int(root.winfo_screenwidth()/2 - windowWidth/2)
    positionDown = int(root.winfo_screenheight()/2 - windowHeight/2)
    child.geometry("+{}+{}".format(positionRight, positionDown))
    msglabel=ttk.Label(child,text = Message,anchor='center')
    msglabel.grid(row=0,columnspan=2, sticky='nwes')
    def exit():
        root.destroy()
    def play_again():
        child.destroy()
        root.deiconify()
        player_1.clear()
        player_2.clear()
        start_button.configure(state='disabled')
    ttk.Button(child,text='Play Again',command=play_again).grid(row=1,sticky='nwes')
    ttk.Button(child,text='Exit',command=exit).grid(row=1,column=1,sticky='nwes')
    child.protocol("WM_DELETE_WINDOW", exit)

def counter_initiallize():
    global counter
    if v.get()==1:
        counter=0
        start_button.configure(state="enabled")
    elif v.get()==2:
        counter=1
        start_button.configure(state="enabled")

def start_game():
    root.withdraw()
    child = Toplevel(root) 
    child.title('Tic tac Toe')
    child.resizable(False,False)
    windowWidth = root.winfo_reqwidth()
    windowHeight = root.winfo_reqheight()
    positionRight = int(root.winfo_screenwidth()/2 - windowWidth/2)
    positionDown = int(root.winfo_screenheight()/2 - windowHeight/2)
    child.geometry("+{}+{}".format(positionRight, positionDown))

    label1 = ttk.Label(child,text='PLAYER1: X',font='bold 10',justify='center')
    label1.grid(row=0,column=0)
    label2 = ttk.Label(child,text='PLAYER2: O',font='bold 10',justify='center')
    label2.grid(row=0,column=2)

    Button1 = ttk.Button(child,command=lambda : assign_value(1))
    Button1.grid(row=1,column=0, ipady=10, ipadx=1,sticky='nwes')
    Button2 = ttk.Button(child,command=lambda : assign_value(2))
    Button2.grid(row=1,column=1, ipady=10, ipadx=1,sticky='nwes')
    Button3 = ttk.Button(child,command=lambda : assign_value(3))
    Button3.grid(row=1,column=2, ipady=10, ipadx=1,sticky='nwes')

    Button4 = ttk.Button(child,command=lambda : assign_value(4))
    Button4.grid(row=2,column=0, ipady=10, ipadx=1,sticky='nwes')
    Button5 = ttk.Button(child,command=lambda : assign_value(5))
    Button5.grid(row=2,column=1, ipady=10, ipadx=1,sticky='nwes')
    Button6 = ttk.Button(child,command=lambda : assign_value(6))
    Button6.grid(row=2,column=2, ipady=10, ipadx=1,sticky='nwes')

    Button7 = ttk.Button(child,command=lambda : assign_value(7))
    Button7.grid(row=3,column=0, ipady=10, ipadx=1,sticky='nwes')
    Button8 = ttk.Button(child,command=lambda : assign_value(8))
    Button8.grid(row=3,column=1, ipady=10, ipadx=1,sticky='nwes')
    Button9 = ttk.Button(child,command=lambda : assign_value(9))
    Button9.grid(row=3,column=2, ipady=10, ipadx=1,sticky='nwes')

    def assign_value(number):
        global value,counter,player_1,player_2
        from itertools import permutations
        set1=permutations([1,2,3])
        set2=permutations([3,5,7])
        set3=permutations([1,5,9])
        set4=permutations([4,5,6])
        set5=permutations([7,8,9])
        set6=permutations([1,4,7])
        set7=permutations([2,5,8])
        set8=permutations([3,6,9])
        combinations=[set1,set2,set3,set4,set5,set6,set7,set8]

        if number == 1:
            if counter%2==0:
                value='X'
                player_1.append(number)
            else:
                value = 'O'
                player_2.append(number)
            Button1.config(text=value)
            Button1.configure(state='disabled')
            counter+=1
        if number == 2:
            if counter%2==0:
                value='X'
                player_1.append(number)
            else:
                value = 'O'
                player_2.append(number)
            Button2.config(text=value)
            Button2.configure(state='disabled')
            counter+=1
        if number == 3:
            if counter%2==0:
                value='X'
                player_1.append(number)
            else:
                value = 'O'
                player_2.append(number)
            Button3.config(text=value)
            Button3.configure(state='disabled')
            counter+=1
        if number == 4:
            if counter%2==0:
                value='X'
                player_1.append(number)
            else:
                value = 'O'
                player_2.append(number)
            Button4.config(text=value)
            Button4.configure(state='disabled')
            counter+=1
        if number == 5:
            if counter%2==0:
                value='X'
                player_1.append(number)
            else:
                value = 'O'
                player_2.append(number)
            Button5.config(text=value)
            Button5.configure(state='disabled')
            counter+=1
        if number == 6:
            if counter%2==0:
                value='X'
                player_1.append(number)
            else:
                value = 'O'
                player_2.append(number)
            Button6.config(text=value)
            Button6.configure(state='disabled')
            counter+=1
        if number == 7:
            if counter%2==0:
                value='X'
                player_1.append(number)
            else:
                value = 'O'
                player_2.append(number)
            Button7.config(text=value)
            Button7.configure(state='disabled')
            counter+=1
        if number == 8:
            if counter%2==0:
                value='X'
                player_1.append(number)
            else:
                value = 'O'
                player_2.append(number)
            Button8.config(text=value)
            Button8.configure(state='disabled')
            counter+=1
        if number == 9:
            if counter%2==0:
                value='X'
                player_1.append(number)
            else:
                value = 'O'
                player_2.append(number)
            Button9.config(text=value)
            Button9.configure(state='disabled')
            counter+=1

        for i in combinations:
            for j in list(i):
                player_1_victory=all(elements in player_1 for elements in j)
                player_2_victory=all(elements in player_2 for elements in j)
                if player_1_victory:
                    CustomMsgBox("Player 1 Wins")
                    child.destroy()
                    break
                elif player_2_victory:
                    CustomMsgBox("Player 2 Wins")
                    child.destroy()
                    break
        
        try:
            if(str(Button1['state'])=='disabled' and
            str(Button2['state'])=='disabled' and
            str(Button3['state'])=='disabled' and
            str(Button4['state'])=='disabled' and
            str(Button5['state'])=='disabled' and
            str(Button6['state'])=='disabled' and
            str(Button7['state'])=='disabled' and
            str(Button8['state'])=='disabled' and
            str(Button9['state'])=='disabled'):
                CustomMsgBox("The match is a draw")
                child.destroy()
        except:
            pass

root = ThemedTk()
root.title("Let's Play")
root.resizable(False,False)
v=IntVar()
label_1 = ttk.Label(root,text="Who Will Start ?",anchor='center')
label_1.grid(row=0,columnspan=2,sticky='nwes')
option_1 = ttk.Radiobutton(root,text="Player 1",variable=v,value=1,command=counter_initiallize)
option_1.grid(row=1,column=0, ipady=2, ipadx=2,sticky='nwes')
option_2 = ttk.Radiobutton(root,text="Player 2",variable=v,value=2,command=counter_initiallize)
option_2.grid(row=1,column=1, ipady=2, ipadx=2,sticky='nwes')
start_button=ttk.Button(root,text="Start Game",command=start_game)
start_button.configure(state="disabled")
start_button.grid(row=2,columnspan=2, ipady=2, ipadx=2,sticky='nwes')

root.mainloop()