"""Coded by Priyam Shah and Aliasgar Haji"""

import play_scraper
import xlwt
category=["ANDROID_WEAR","BOOKS_AND_REFERENCE","BUSINESS","COMICS","COMMUNICATION","EDUCATION","ENTERTAINMENT","FINANCE","HEALTH_AND_FITNESS","LIBRARIES_AND_DEMO","LIFESTYLE",
         "APP_WALLPAPER","MEDIA_AND_VIDEO","TOOLS","TRANSPORTATION","MEDICAL","MUSIC_AND_AUDIO","NEWS_AND_MAGAZINES","PERSONALIZATION","PHOTOGRAPHY","PRODUCTIVITY","SHOPPING","SOCIAL","SPORTS",
         "TRAVEL_AND_LOCAL","WEATHER","APP_WIDGETS","GAME_ACTION","GAME_ADVENTURE","GAME_ARCADE","GAME_BOARD","GAME_CARD","GAME_CASINO",
         "GAME_CASUAL","GAME_EDUCATIONAL","GAME_MUSIC","GAME_PUZZLE","GAME_RACING","GAME_ROLE_PLAYING","GAME_SIMULATION","GAME_SPORTS","GAME_STRATEGY",
         "GAME_TRIVIA","GAME_WORD","FAMILY","FAMILY_ACTION","FAMILY_BRAINGAMES","FAMILY_CREATE","FAMILY_EDUCATION","FAMILY_MUSICVIDEO","FAMILY_PRETEND"]
collect=['NEW_FREE','NEW_PAID','TOP_FREE','TOP_PAID','TOP_GROSSING','TRENDING']
k=1
l=0
wb=xlwt.Workbook()
for m in category:
    k=1
    l=0
    ws=wb.add_sheet(m)
    ws.write(0,0,"App_Id")
    ws.write(0,1,"URL")
    ws.write(0,2,"Icon")
    ws.write(0,3,"Title")
    ws.write(0,4,"Developer")
    ws.write(0,5,"Developer_id")
    ws.write(0,6,"Description")
    ws.write(0,7,"Score")
    ws.write(0,8,"Price")
    ws.write(0,9,"Free")



    for u in collect:
        try:
            app=play_scraper.collection(
                collection=u,
                category=m,
                results=50,
                page=0)
            for i in app:
                
        


                for j in i.values():
                    ws.write(k,l,j)
                    l=l+1
                k=k+1
                l=0
        
        except:
            continue;
        wb.save("PlayStore.xls")
    
print("finally created")
