# -*- coding: utf-8 -*-
"""
Created on Sun Jul 17 14:38:02 2022

@author: johnm
"""

import bs4, requests, re, random, time, openpyxl

castlist=[]
castlinks=[]
footheight=[]
inchesheight=[]

DIRECTORY = "none"

#This function pulls the castmembers and links and finds height in feet and inches (Europoors on suicide watch)
def bottrickster():
    global footheight
    global inchesheight
    castlinkslen=len(castlinks)
    
    #If you want to only run through the first few cast members, change next line to that number. Remember to change back to castlinkslen. 
    for i in range(0,castlinkslen,1):
        sleeptime = random.randint(3,7)
        print("waiting " + str(sleeptime) +" seconds...")     
        time.sleep(int(sleeptime))
        res = requests.get('https://www.imdb.com/name/'+str(castlinks[i])+'/?ref_=ttfc_fc_cl_t1')
        res.raise_for_status()
        soup = bs4.BeautifulSoup(res.text, 'html.parser')
        detailselector=soup.select('.see-more')
        details=str(detailselector)
   
        heightregex = re.compile(r'<h4 class="inline">Height:</h4>\n(.*)?" ')
        heightlist = heightregex.findall(details)
        heightlist=str(heightlist)
        heightlistfixerregex = re.compile(r'\d')
        castheight=heightlistfixerregex.findall(heightlist)
        

        if len(castheight)>=1:
            footheight.append(castheight[0])
        else:
            footheight.append("NULL")
        
        if len(castheight)>=1:
            if castheight[4]!=str("0"):
                clownass=castheight[2]  
            else:
                if castheight[3]==str("0"):
                    clownass=str("10")
                else:
                    clownass=str("11")
        else:
            clownass=str("NULL")  
        inchesheight.append(clownass)
        
        #Prints each height in feet and inches. I mainly use it to keep track of program running, vs any real purpose
        print(str(footheight[i])+" ft "+str(inchesheight[i])+" in ")
    return footheight
    return inchesheight


#The program starts running HERE:

if DIRECTORY == str("none"):
    DIRECTORY = input("WARNING: No directory selected!\nPaste directory here, or add directory into \"DIRECTORY\" variable in code.\nNOTE: remember to use double-backslashes when pasting")
else:
    print("directory detected")

thisthing = input("What is the movie number of the IMDB movie you want to search?")



#This code pulls the castlist and turns it into a string so that you can read it on your own machine
res = requests.get('https://www.imdb.com/title/'+thisthing+'/fullcredits?ref_=tt_ov_st_sm')
res.raise_for_status()
soup = bs4.BeautifulSoup(res.text, 'html.parser')
castlistselector = soup.select('.cast_list a')
castlist2 = str(castlistselector)

#This code creates a list of the cast members in the movie
castregex = re.compile(r'title="(.*)?" ')
castlist = castregex.findall(castlist2)

#This code creates a list of the links to the actors in the movies' personal imdb pages
castlinkregex = re.compile(r'<a href="/name/(.*)?/"><img')
castlinks = castlinkregex.findall(castlist2)


bottrickster()

#This part puts all the important info onto an excel sheet
wb = openpyxl.Workbook()
sheet = wb.active

sheet['A1'] = "Actors' Names"
sheet['B1'] = "Feet"
sheet['C1'] = "Inches"
for i in range(0,len(footheight)):
   num = i+2
   sheet['A'+ str(num)] = castlist[i]
   if footheight[i]==str("NULL"):
      sheet['B'+ str(num)] = str("NULL")
      sheet['C'+ str(num)] = str("NULL")
   else:
      sheet['B'+ str(num)] = int(footheight[i])
      sheet['C'+ str(num)] = int(inchesheight[i])
      
wb.save(str(DIRECTORY)+'\\Movie casts\' heights.xlsx')

input("All finished. Press ENTER Key to exit.")