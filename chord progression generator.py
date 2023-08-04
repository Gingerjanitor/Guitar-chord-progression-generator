import openpyxl as pyxl
import random
from collections import defaultdict
import pprint
###Goal: make a simple chord progression generator that will ask you the key you want to play in (major or minor, too).
###it will have a dictionary for each key. The dictionary will contain the key:name; I: chord, II: chord; etc etc.
#dictionaries will be merged to form a data structure.
#from there, the user's entry will be processed:- search the dictionaries for whatever matches the key. Then, print all the chords of the key in the
#circle of 5ths box (use that same method as prock paper scissors interface here).
# All the chord names.
#"How about you play with these chords: figure out your own ordering!
#maybe this progression? Randomly shuffle those four chords and spit it out. What mood does this progression have?

#data will have to be structured as {key:{chord:name , key2:{chord:name} etc etc}
#right now the hope is to import cells from an excel file. idk how that will work. Only one way to find out....

#*excel file lives in:
#"C:\Users\Matt0\Documents\python learning\Automate the boring stuff\musical chord generator\Chord  dataset.xlsx"?

key={}
merged = {}

letters=["A","B","C","D","E","F","G", "H","I"]
def importkeys():
    workbook=pyxl.load_workbook(filename="C:/Users/Matt0/Documents/python learning/Automate the boring stuff/musical chord generator/Chord  dataset.xlsx")
    #wb=workbook(loc)
    #ws=wb.active
    #wb=workbook()
    #ws=wb.active
    #products_sheet = workbook["Products"]
    #sales_sheet = workbook["Company Sales"]
    chordlist=workbook["chordlist"]

    keyspot = 2
    rowspace = 1
    done = 0
    #finallist=[]
    count=0
    while done==0:
        chord = []
        interval = []

        letter = 1
        paired = []
        key = 0

        for loop in range(len(letters)):
            try:
                dangervalue=str((chordlist[f"{letters[letter-1]}{rowspace}"].value))
                if dangervalue.lower()=="none":
                    done+=1
                    break
                else:
                    interval.append(chordlist[f"{letters[letter-1]}{rowspace}"].value)
                    chord.append(chordlist[f"{letters[letter-1]}{rowspace+1}"].value)
                    letter+=1

            except IndexError:
                break
            key=(chordlist[f"J{keyspot}"].value)
        if done!=1:
            paired=dict(zip(interval,chord))
            merged[key]=paired
            #finallist.append(merged)
            keyspot += 2
            rowspace += 2
            count+=1

    #pprint.pprint(merged)
    return merged


def showtable(key,majmin):
    majinter=["IV","I","V","ii","vi","iii"]
    mininter=["VI","III","VII","iv","i","v"]
    interval=[]
    if majmin==1: #major
        interval=majinter
    if majmin==2: #minor
        interval=mininter
    print(f"""\n  {interval[0]}: {chords[key][interval[0]]}  |  {interval[1]}: {chords[key][interval[1]]}  |  {interval[2]}: {chords[key][interval[2]]}""")
    print("------------------------------------------")
    print(f"""  {interval[3]}: {chords[key][interval[3]]}  |  {interval[4]}: {chords[key][interval[4]]}  |  {interval[5]}: {chords[key][interval[5]]}""")
    print(f"""\nThere is also the diminished chord, {chords[key]["dimin"]}""")
    print(f"""The relative of this key is {chords[key]["Relative"]}""")
    return interval
   #print(board['mid-l'] + " |" + board['mid-m']+ " |" + board["mid-r"])

def genprogression(majmin, complex):
    if majmin==1:
        validchords=[0,1,2,3,4,5]
        progression=[0]
    elif majmin==2:
        validchords=[0,1,2,3,4,5]
        progression=[0]
    if complex==1:
        difficulty=random.randint(1,3)
        letsplay=random.choices(validchords,weights=[16.66,16.66,10,23.32,16.66,16.66],k=difficulty)
        finalchords=progression+letsplay
    elif complex==2:
        difficulty=random.randint(4,6)
        letsplay=random.choices(validchords,weights=[16.66,16.66,10,23.32,16.66,16.66],k=difficulty)
        finalchords=progression+letsplay
    elif complex==3:
        difficulty = random.randint(7,9)
        letsplay=random.choices(validchords,weights=[16.66,16.66,10,23.32,16.66,16.66],k=difficulty)
        finalchords=progression+letsplay
    return finalchords

def displaysuggestsions(finalchords, key, majmin):
    simpinterval=[]
    if majmin==1:
        simpinterval=["I","ii","iii","IV","V","vi"]
    elif majmin==2:
        simpinterval=["i","III","iv","v","VI","VII"]
    print(f"The progression you ought to try is:")
    for loop in range(len(finalchords)):
        print (f"{simpinterval[finalchords[loop]]}: {chords[key][simpinterval[finalchords[loop]]]}")





##TODO-ok, so how can I make it so that it comes up with a random chord progression? Perhaps draw a random set of numbers from that list I wrote above?

acceptmaj=["Cmaj","Dmaj","Emaj", "Fmaj", "Gmaj", "Amaj","Bmaj","Gbmaj","Dbmaj","Abmaj","Ebmaj","Bbmaj"]
acceptmin=["Amin","Emin","Bmin", "F#min","C#min","G#min","Ebmin", "Bbmin", "Fmin","Cmin","Gmin", "Dmin"]
for loop in range(len(acceptmaj)):
    print(f"{acceptmaj[loop]}     | {acceptmin[loop]} ")

key=input("Hi there! Pick a key from above and I'll try my best to come up with a chord progression. \nFor best results enter it exactly as in the list.")

key=key.capitalize()
if key.capitalize() not in ["Cmaj","Dmaj","Emaj", "Fmaj", "Gmaj", "Amaj","Bmaj","Gbmaj","Dbmaj","Abmaj","Ebmaj","Bbmaj" ,"Amin","Emin","Bmin", "F#min","C#min","G#min","Ebmin", "Bbmin", "Fmin","Cmin","Gmin", "Dmin"]:
    input("I don't think I know that key. Try formatting like Cmaj, Amin, Eflatmin, or F#min")

#majmin=int(input("\nJust to confirm, press 1 if that was major, 2 if minor"))
if key.capitalize() in ["Cmaj","Dmaj","Emaj", "Fmaj", "Gmaj", "Amaj","Bmaj","Gbmaj","Dbmaj","Abmaj","Ebmaj","Bbmaj"]:
    majmin=1
elif key.capitalize() in ["Amin","Emin","Bmin", "F#min","C#min","G#min","Ebmin", "Bbmin", "Fmin","Cmin","Gmin", "Dmin"]:
    majmin=2



chords=importkeys()

#print(finallist)
#print(finallist['Amin']['Relative'])


interval=showtable(key, majmin)

complex=int(input("\n\nHow complex should the progression be? Simple (2-4 chords), Moderate (5-7 chords), complex (8+ chords)?\n For simple, enter 1, 2=moderate, 3=complex"))
while complex not in [1,2,3]:
    complex=input("\n\n THat wasn't a valid entry. For simple, enter 1, 2=moderate, 3=complex")
finalchords=genprogression(majmin, complex)

displaysuggestsions(finalchords, key, majmin)

input("\nthat's it for now. Planning to add GUI of some sort to study that idea\n \nPress a key to kill.")
