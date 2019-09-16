import pyodbc, re, os, sys, time
import tkinter as tk
from tkinter import messagebox
from tkinter import BOTH, END, LEFT
from tkinter import filedialog
from tkinter import *

txtInput = ''
text =''
entry = ''
string = ''
leng = 0
readin = ''
highestNo = 0
highestWord = ''
badWord = ''
words = [''] * 2500
weeks = [''] * 2500
counter = 0
week = 0
word = ''
tryWord = ''
splitReadin = '' #THESE ARE ALL THE VARIABLES THAT HAVE TO BE INITIALISED BEFORE BEING ABLE TO USE THEM IN THE SCRIPT
outputPrint = ''
file = ''
dirBtn = ''
highFreqWords = [''] * 2500
highFreqWordsBook = [''] * 2500
trickyWords = [''] * 2500
trickyWordsBook = [''] * 2500
bookVar = ''
firstTime = False
finalOutput = ''
frame2 = ''
root2 = ''
report = ''


def returnTrickyWords(strSplit, crsr):
    global trickyWords
    global trickyWordsBook

    counter = 0
    returnStr = ''

    crsr.execute('SELECT T_WordDatabase.Word, T_WordDatabase.Tricky, T_WordDatabase.Book FROM T_WordDatabase WHERE (((T_WordDatabase.Tricky)=1));') #SETTING UP THE DATABASE QUERY

    for row in crsr.fetchall(): #WRITES THE DATABASE INTO OUR OWN ARRAYS
        trickyWords[counter] = row[0]
        trickyWordsBook[counter] = row[2]
        counter = counter + 1

    strSplit = list(set(strSplit))
    bookNumber = int(getOption())
    for i in range(len(strSplit)):
        if strSplit[i] in trickyWords:
            if trickyWordsBook[trickyWords.index(strSplit[i])] == bookNumber:
                if strSplit[i] != '':
                    returnStr = returnStr + strSplit[i] + ', '

    if returnStr == '':
        return ''

    if returnStr != '':
        return returnStr[:-2]


def returnHighFreqWords(strSplit, crsr):
    global highFreqWords
    global highFreqWordsBook

    counter = 0
    returnStr = ''

    crsr.execute('SELECT T_WordDatabase.Word, T_WordDatabase.HighFrequency, T_WordDatabase.Book, T_WordDatabase.Tricky FROM T_WordDatabase WHERE (((T_WordDatabase.HighFrequency)=1) AND ((T_WordDatabase.Tricky) Is Null));') #SETTING UP THE DATABASE QUERY

    for row in crsr.fetchall(): #WRITES THE DATABASE INTO OUR OWN ARRAYS
        highFreqWords[counter] = row[0]
        highFreqWordsBook[counter] = row[2]
        counter = counter + 1

    strSplit = list(set(strSplit))
    bookNumber = int(getOption())
    for i in range(len(strSplit)):
        if strSplit[i] in highFreqWords:
            if highFreqWordsBook[highFreqWords.index(strSplit[i])] == bookNumber:
                if strSplit[i] != '':
                    returnStr = returnStr + strSplit[i] + ', '

    if returnStr == '':
        return ''

    if returnStr != '':
        return returnStr[:-2]


def storyChecker(string):
    global txtInput
    global text
    global entry
    global leng
    global readin
    global highestNo
    global highestWord
    global badWord
    global words
    global weeks
    global counter
    global week
    global word
    global tryWord
    global splitReadin
    global outputPrint
    global file
    global firstTime
    global finalOutput

    highestWord = ''
    highestNo = 0
    outputMessage = ''
    counter = 0
    conn_str = ( #FROM HERE
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=' + file + ';' #FILE LOCATION OF THE DATABASE
        )
    try:
        cnxn = pyodbc.connect(conn_str)
        crsr = cnxn.cursor()
    except:
        messagebox.showinfo("ERROR", "You didn't select a database")


    crsr.execute('SELECT T_WordDatabase.Word, T_WordDatabase.[Week no] FROM T_WordDatabase ORDER BY T_WordDatabase.Word;') #SETTING UP THE DATABASE QUERY

    string1 = ''

    for i in range(len(string)):
        if string[i] != '\n':
            string1 = string1 + string[i]
        if string[i] == '\n':
            if string[i-1] != '\n':
                string1 = string1 + " "

    puncExceptions = [''] * 100
    counter10 = 0
    string = string1.lower()
    strSplit = string.split(' ')
    strSplit = list(set(strSplit))

    for i in range(len(strSplit)):
        try:

            for i in range(len(strSplit)):
                var = re.findall(r"\w+"+r"'"+r"\w+", strSplit[i])
                if var == []:
                    for k in strSplit[i].split("\n"):
                        strSplit[i] = re.sub(r"[^a-zA-Z0-9]+", '', k)

                if var != []:
                    strSplit[i] = var[0]

        except:
            ye = 'man'

    try:
        for i in range(len(strSplit)):
            if strSplit[i] == '':
                del strSplit[i]
    except:
        ye = 'man'
    # strSplit.pop(0) #REMOVES THE FIRST INDEXC IN THE ARRAY AS IT IS BLANK

    if strSplit[len(strSplit) - 1] == '':
        strSplit.pop(len(strSplit) - 1) #THE LAST INDEX IN THE ARRAY IS BLANK, THIS REMOVES IT


    for row in crsr.fetchall(): #WRITES THE DATABASE INTO OUR OWN ARRAYS
        words[counter] = row[0]
        weeks[counter] = row[1]
        counter = counter + 1


    for i in range (len(strSplit)): #LOOPS SAME AMOUNT OF TIMES AS THERE ARE WORDS IN THE STORY
        if strSplit[i] in words: #CHECKING EACH WORD IN THE STORY AGAINST OUR ARRAY OF WORDS FROM THE DATABASE
            week = weeks[words.index(strSplit[i])] #CHECKS WHAT INDEX A PARTICULAR WEEK IS AT IN THE WEKS ARRAY AND REWRITES THE WEEK VARIABLE
            word = words[words.index(strSplit[i])] #SAME GOES FOR THIS, ONLY WORDS INSTEAD

            # print(word + " " + str(week))
            outputMessage = outputMessage + word + "\t\t" + str(week) + "\n"

            if week > highestNo: #A QUEST TO FIND THE HIGHEST WEEK, IF THERE IS A HIGHER WEEK, THE HIGHESTNO AND HIGHESTWORD VARIABLES GET REWRITTEN
                highestNo = week
                highestWord = word


        if strSplit[i] not in words: #IF A WORD IN THE STORY IS NOT IN THE DATABSE
            tryWord = strSplit[i]
            tryWord = str(tryWord[0]).upper() + tryWord[1:len(tryWord)] #CONVERTS THE FIRST LETTER OF THE WORD TO AN UPPER CASE E.G. cat to Cat
            if tryWord in words: #CHECKS AGAIN TO SEE IF IT WAS A CASE ERROR
                week = weeks[words.index(tryWord)]
                word = words[words.index(tryWord)]
                # print(word + " " + str(week))
                outputMessage = outputMessage + word + "\t\t" + str(week) + "\n"
                if week > highestNo: #FINDS THE HIGHEST AGAIN
                    highestNo = week
                    highestWord = word
            if tryWord not in words: #IF THE WORD STILL ISN'T IN WORDS, IT'S HIGHLY LIKELY THE WORD IS NOT IN THE DATABASE
                # print(str(tryWord) + " is not in database")
                outputMessage = outputMessage + str(tryWord) + " is not in the database" + "\n"

    counter = 0
    highFreq = returnHighFreqWords(strSplit, crsr)
    highFreq = "HF100: " + highFreq

    tricky = returnTrickyWords(strSplit, crsr)
    tricky = "TW: " + tricky

    findNum = []
    inOrder = []
    leftOver = []
    placeHolder = 0
    var20 = outputMessage.split("\n")
    var20.remove('')
    # print(var20)

    for i in range (len(var20)):
        try:
            findNum.append(int(re.findall(r"\t\t(\d+)",str(var20[i]))[0]))
        except:
            leftOver.append(var20[i])
        # print(int(re.findall(r"\t\t(\d+)",str(var20[i]))[0]))

    # print(findNum)
    findNum.sort(reverse=True)
    # if findNum[0] == '':
    #     findNum.pop(0)


    # print(var20)
    for i in range(len(findNum)):
        for j in range(len(var20)):
            try:
                if re.search(str(findNum[i]), var20[j]) != None and var20[j] not in inOrder:
                    inOrder.append(var20[j])

            except:
                ye = "man"

    outputMessage = ''
    leftOverMessage = ''
    for i in range(len(inOrder)):
        outputMessage = outputMessage + inOrder[i] + "\n"

    for i in range(len(leftOver)):
        leftOverMessage = leftOverMessage + leftOver[i] + "\n"



    finalOutput = "The maximum week for this story is Week " + str(highestNo) + " caused by the word '" + highestWord + "'\n" + highFreq + "\n" + tricky + "\n\n" + outputMessage + "\n" + leftOverMessage

def main():
    global txtInput
    global text
    global entry
    global highestNo
    global outputPrint
    global entryText
    global dirBtn
    global bookVar
    global firstTime
    global finalOutput
    global root2
    global frame2
    global report

    counter = 0
    outputPrint = ''
    root = tk.Tk()
    root.title("Story Week Checker")
    w = 1000
    h = 800
    x = 50
    y = 100
    root.geometry("%dx%d+%d+%d" % (w, h, x, y))
    frame = tk.Frame(root, bg='white')
    frame.pack(fill='both', expand='yes')
    frame.update()
    label = tk.Label(frame, text="Insert Story: ", font=("Century Gothic", 15))
    label.place(x=20, y=30)

    button = tk.Button(frame, text="Check Story!", bg='white', font=("Century Gothic", 15), command=printFunc)
    button.pack()
    frame.update()
    button.place(x=20, y=frame.winfo_height()-100)

    scrollbar = tk.Scrollbar(frame)
    entry = tk.Text(frame, font=("Century Gothic", 20), width=50, height=25,  yscrollcommand=scrollbar.set)
    entry.pack()
    entry.place(x=180, y=30)

    dirBtn = tk.Button(frame, text="Choose Database", bg='white', font=("Century Gothic", 12), command=openFunc)
    dirBtn.pack()
    frame.update()
    dirBtn.place(x=15, y=100)

    bookVar = tk.StringVar(root)
    choices = ['1','2','3','4','5','6','7','8','9']
    bookVar.set('1')
    popupMenu = OptionMenu(frame, bookVar, *choices)
    popupMenu.place(x=65, y=220)

    label2 = tk.Label(frame, text="Choose a book \n number:", font=("Century Gothic", 14))
    label2.place(x=15, y=160)

    root2 = tk.Tk()
    frame2 = tk.Frame(root2, bg='white', takefocus=TRUE)
    frame2.pack(fill='both', expand='yes')
    root2.title("Check Report")
    w = 800
    h = 800
    x = 50
    y = 100
    root2.geometry("%dx%d+%d+%d" % (w, h, x, y))
    frame2.update()
    scrollbar = tk.Scrollbar(frame2)
    report = tk.Text(frame2, font=("Century Gothic", 15), width=68, height=32,  yscrollcommand=scrollbar.set)
    report.pack()
    report.place(x=20, y=10)

    root.mainloop()

def openFunc():
    global file
    global dirBtn
    file1 = ''
    file = filedialog.askopenfile(mode='rb',title='Choose a file')
    if file != None:
        file = str(file)[26:len(str(file))-2]
        for i in range(len(file)):
            if file[i] == '/':
                file1 = file1 + '\\'
            if file[i] != '/':
                file1 = file1 + file[i]
    file = file1

def printFunc():
    global entry
    global entryText
    global firstTime
    global report
    global frame2
    string = entry.get(1.0, END)
    storyChecker(string)
    firstTime = True
    frame2.update()
    report.delete(1.0, END)
    frame2.update()
    report.insert(INSERT, finalOutput)


def getOption():
    global bookVar
    return bookVar.get()

if __name__ == '__main__':
    main()



# can't
# scrolling on report box
# ascending weeks
# main report conclusion at the top
