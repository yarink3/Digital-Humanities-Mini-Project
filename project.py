import os
from openpyxl import Workbook
import pandas as pd
import json
import time
from selenium import webdriver
from geopy.geocoders import Nominatim
from selenium.webdriver.chrome.options import Options
path_to_driver="./chromedriver"

def create_corpus():
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"]="name"
    sheet["B1"]="short_info"
    sheet["C1"]="studies"
    sheet["D1"]="teaching"
    sheet["E1"]="awards"
    sheet["F1"]="num_of_awards"
    sheet["G1"]="statues_outside"

    current_artist=2;

    option = webdriver.ChromeOptions()
    option.add_argument('headless')
    driver = webdriver.Chrome(path_to_driver, options=option)

    urls= open('list_of_artists_urls.txt','r', encoding="utf-8")
    found=True
    for url in urls:
        driver.get(url)
        info=driver.find_elements_by_class_name("left")
        for x in info:
            headers=x.find_elements_by_tag_name("strong")
            paragraphs=x.find_elements_by_tag_name("p")

            headers_list=[]
            paragraphs_list=[]
            for header in headers:
                headers_list.append(header.text)
            for par in paragraphs:
                paragraphs_list.append(par.text)
            # break

            artist_name=headers_list[0]
            sheet["A" + str(current_artist)] =artist_name
            print("working on : "+artist_name)
            has_short_info=True
            if(len(paragraphs_list)+2==len(headers_list)):
                par_index=0
                has_short_info=False

            else:
                par_index=1
                if(has_short_info):
                    arrtist_short_info=paragraphs_list[0]
                    sheet["B" + str(current_artist)] =arrtist_short_info


                for header in range(2,len(headers_list)):
                    if(str(headers_list[header])=='לימודים'):
                        sheet["C"+str(current_artist)]=paragraphs_list[par_index]
                        par_index=par_index+1
                    elif (str(headers_list[header]) == 'הוראה'):
                        sheet["D" + str(current_artist)] = paragraphs_list[par_index ]
                        par_index=par_index+1
                    elif (str(headers_list[header]) == 'פרסים'):
                        sheet["E" + str(current_artist)] = paragraphs_list[par_index ]
                        awards_counter=paragraphs_list[par_index ].count('\n') + 1
                        sheet["F" + str(current_artist)] = str(awards_counter)
                        par_index=par_index+1

                    elif (str(headers_list[header]) == 'פסלים במרחב הציבורי'):
                        sheet["G" + str(current_artist)] = paragraphs_list[par_index ]
                        par_index = par_index + 1
        current_artist=current_artist+1
    driver.quit()
    workbook.save(filename="artists_corpus_new_with_info.xlsx")


def create_file_of_artists_names_and_urls():
    driver=webdriver.Chrome(path_to_driver)
    url = "https://museum.imj.org.il/artcenter/newsite/he/?list="
    hebrewLetters=[ "א", "ב", "ג", "ד", "ה", "ו", "ז", "ח", "ט", "י", "כ", "ל", "מ", "נ", "ס", "ע", "פ", "צ", "ק", "ר", "ש",  "ת" ]
    with open('list_of_artists_urls.txt', 'a', encoding="utf-8") as urls:
     with open('list_of_artists_names.txt', 'a', encoding="utf-8") as artists_names_file:
        for letter in hebrewLetters:
            print("started letter: "+letter)
            driver.get(url+letter)
            tableHeaders = driver.find_elements_by_class_name("list_of_artists")
            for artist in tableHeaders:
               x=artist.find_elements_by_tag_name("li")
               for y in x:
                artists_names_file.write(y.text+"\n")
                urls.write(get_artist_url(y.text,letter)+"\n")

            driver.quit()


        urls.close()
        artists_names_file.close()
        print("finished")

        driver.quit()

def add_20(name):
    for i in range(len(name)):
        if(name[i]==' '):
            name=name[0:i]+"%20"+name[i+1:len(name)]

    return name


def get_artist_url(name,letter):
    start="https://museum.imj.org.il/artcenter/newsite/he/?artist="
    new_name=add_20(name)
    url=start+new_name

    return url


def write_file(path, data):
    file=open(path,'w')
    file.write(data+'\n')
    file.close()


def create_project_dir(directory):
    if not os.path.exists(directory):
        print("Creating project: "+directory)
        os.makedirs(directory)

def create_data_files(project_name,base_url):
    queue=project_name + "/queue.txt"
    crawled=project_name + "/crawled.txt"
    if not os.path.isfile(queue):
        write_file(queue,base_url)
    if not os.path.isfile(crawled):
        write_file(crawled,'')

def append_to_file(path,data):
        file= open(path,'a', encoding="utf-8")
        # print("link: "+data)
        # print("55-59: "+data[55:60])
        # new=change_amper(data)
        file.write(data+'\n')
        # file.close()

def delete_file_content(path):
    with open(path, 'w') :
        pass


def file_to_set(file_name):
    results=set()
    with open(file_name, 'rt') as file:
        for line in file:
            results.add(line)
    return results

def set_to_file(links,file):
    for link in links:
        append_to_file(file,link)

def nlp_search(str):
    url = "https://hebrew-nlp.co.il/service/ner/Cities"
    try:
        options = Options()
        options.add_argument("--headless")
        driver = webdriver.Chrome(options=options)

        options = webdriver.ChromeOptions()
        options.add_argument('headless')


        driver.get(url)
        # s=requests.session()
        password="AoLM5A0mFqYclnF"


        driver.find_element_by_id("token").send_keys(password)
        driver.find_element_by_id("readable")
        readable = driver.find_element_by_id('readable')
        for option in readable.find_elements_by_tag_name('option'):
            if option.text == 'מפורט':
                option.click() # select() in earlier versions of webdriver
                break
        driver.find_element_by_id("text").send_keys(str)
        driver.find_element_by_xpath("/html/body/main/section/div/div[2]/div[2]/div/div[6]/button").click()
        time.sleep(2)
        ret=driver.find_element_by_xpath("/html/body/main/section/div/div[2]/div[4]/div/pre")

        # file.write(ret.text+",\n")
        print(ret.text)
        x=ret.text
        y=json.loads(x)
        driver.close()
        return y

    except:
        return None





def get_country(city):
    try:
        geolocator = Nominatim(timeout=3,user_agent="geoapiExercises")

        location = geolocator.geocode(city,language='he')

        loc_dict = location.raw
        return loc_dict['display_name'].rsplit(',' , 1)[1][1:len(loc_dict['display_name'].rsplit(',' , 1)[1])]
        # print (loc_dict['display_name'].rsplit(',' , 1)[1][1:len(loc_dict['display_name'].rsplit(',' , 1)[1])])
    except:
        return "wasn't found"

def get_area(city):
    try:
        geolocator = Nominatim(timeout=3,user_agent="geoapiExercises")
        location = geolocator.geocode(city,language='he')
        loc_dict = location.raw
        area= loc_dict['display_name'].split(",")
        for i in range (len(area)):
            if(area[i].find("מחוז")!=-1):
                return area[i][1:len(area[i])]
    except:
        return "wasn't found"

def is_center(city):
    if(city==""):
        return False
    return get_area(city) in ["מחוז תל אביב","מחוז המרכז"]

def add_places_of_birth():


    born_places_prefixes = ["נולד","נולדה","יליד","ילידת","גדל", "גדלה","התיישב","התיישבה","חבר קיבוץ", "חברת קיבוץ","ממייסדי","חי ויוצר","חיה ויוצרת", "תושב","תושבת","גר","גרה", "מתגורר","מתגוררת","חי","חיה"]

    df = pd.read_excel('only_infos.xlsx', sheet_name=0)
    df.set_index('short_info',inplace=True)


    with open("json_after_nlp.json", "r", encoding='utf8') as file:
        data = json.load(file)
    israel_general=[]
    city_in_israel=[]
    out_of_israel=[]
    center_lst=[]
    wasnt_found=[]
    for artist_id in range (len(data)):
        found=False
        print(data[str(artist_id)])
        if(data[str(artist_id)]==None):
            israel_general.append("")
            city_in_israel.append("")
            out_of_israel.append("")
            wasnt_found.append("wasn't found")
        else:
            for sentence_id in range(len(data[str(artist_id)])):
                sentence=data[str(artist_id)][sentence_id]
                if(not found):
                    for word_id in range (len(sentence)):
                        # born in a city in israel
                        if "categories" in sentence[word_id].keys() and  "עיר" in sentence[word_id]["categories"] and "token" in sentence[word_id-1].keys() and  sentence[word_id-1]["token"] in born_places_prefixes:
                            print( sentence[word_id-1]["token"]+" ב"+sentence[word_id]["entity"])
                            if(get_country(sentence[word_id]["entity"])=="ישראל"):
                                israel_general.append("")
                                city_in_israel.append(sentence[word_id]["entity"])
                                center_lst.append(is_center(sentence[word_id]["entity"]))
                                out_of_israel.append("")
                                wasnt_found.append("")
                                found=True
                                break
                            else:
                                print(sentence[word_id - 1]["token"] + " ב" + sentence[word_id]["entity"])
                                israel_general.append("")
                                city_in_israel.append("")
                                center_lst.append(False)
                                out_of_israel.append(sentence[word_id]["entity"])
                                wasnt_found.append("")
                                found = True
                                break

                        elif "categories" in sentence[word_id].keys() and  "עיר" in sentence[word_id]["categories"] and word_id>2 and "token" in sentence[word_id-2].keys() and  sentence[word_id-1]["token"] in born_places_prefixes:
                            print( sentence[word_id-2]["token"]+" ב"+sentence[word_id]["entity"])
                            israel_general.append("")
                            city_in_israel.append(sentence[word_id]["entity"])
                            center_lst.append(False)
                            out_of_israel.append("")
                            wasnt_found.append("")
                            found = True
                            break

                        elif "categories" in sentence[word_id].keys() and  "ארץ" in sentence[word_id]["categories"] and sentence[word_id]["entity"]=="ישראל" and "token" in sentence[word_id-1].keys() and  sentence[word_id-1]["token"] in born_places_prefixes:
                            print( sentence[word_id-1]["token"]+" ב"+sentence[word_id]["entity"])
                            israel_general.append("ישראל")
                            city_in_israel.append("")
                            center_lst.append(False)
                            out_of_israel.append("")
                            wasnt_found.append("")
                            found = True
                            break
                        elif "categories" in sentence[word_id].keys() and  "ארץ" in sentence[word_id]["categories"] and sentence[word_id]["entity"]!="ישראל" and "token" in sentence[word_id-1].keys() and  sentence[word_id-1]["token"] in born_places_prefixes:
                            print( sentence[word_id-1]["token"]+" ב"+sentence[word_id]["entity"])
                            israel_general.append("")
                            city_in_israel.append("")
                            center_lst.append(False)
                            out_of_israel.append(sentence[word_id]["entity"])
                            wasnt_found.append("")
                            found = True
                            break
            if (not found):
                israel_general.append("")
                city_in_israel.append("")
                center_lst.append(False)
                out_of_israel.append("")
                wasnt_found.append("wasn't found")



    df["city"]=city_in_israel
    df["israel_general"]=israel_general
    df["out_of_israel"]=out_of_israel
    df["wasnt_found"]=wasnt_found
    df["is_center"]=center_lst

    df.to_excel("cities.xlsx")


def add_galleries_and_hex():
    all_galleries=[]
    all_hexibitions=[]
    options = Options()
    options.add_argument("--headless")
    driver = webdriver.Chrome(options=options)

    options = webdriver.ChromeOptions();
    options.add_argument('headless');
    # driver = webdriver.Chrome(PATHtoChromeDtiver)
    df = pd.read_excel('check_city4.xlsx', sheet_name=0)

    urls = open('list_of_artists_urls.txt', 'r', encoding="utf-8")
    i=0
    for url in urls:
        a=url.find("?artist")
        url_of_galleries=url[0:a]+"gallery/"+url[a:len(url)]
        try:
            driver.get(url_of_galleries)
            num=len(driver.find_elements_by_tag_name("td"))
            print(num)
            all_galleries.append(num)

        # print(num)
        except:
            print("error")
            all_galleries.append(0)


        url_of_hex=url[0:a]+"exhibitions/"+url[a:len(url)]
        try:
            driver.get(url_of_hex)
            num = len(driver.find_elements_by_tag_name("table"))
            print(num)
            all_hexibitions.append(num)


            # print(num)
        except:
            print("error")
            all_galleries.append(0)

        i=i+1
        print(str(i)+ " created already")
        # print(url_of_galleries)
    driver.quit()
    df["num_of_gal"]=all_galleries
    df["num_of_hex"]=all_hexibitions
    df.to_excel("gals_and_hexes.xlsx")


def find_exact(line,word):
    for check in line:
        if(check==word):
            return 1

    return -1

def find_gender():

    female_prefixes = ["ציירת","ישראלית","פסלת","סופרת","היא","למדה","נולדה","הייתה","גורשה","ילידת","החלה","ישראלית","עברה","עבדה"]
    male_prefixes = ["צייר","ישראלי","פסל","סופר","הוא","למד","נולד","היה","גורש","יליד","החל","ישראלית","עברה","עבדה"]
    ret_lst=[]
    df = pd.read_excel('only_infos.xlsx', sheet_name=0)
    columns = len(df.columns)
    rows = len(df)
    i=0
    for column in range(columns):
        for row in range(rows):
            row_str=str(df.loc[row][column])
            if(row_str==""):
                ret_lst.append("unknown")
            else:
                found=False
                for word in female_prefixes:
                    if (find_exact(row_str,word))!= -1 and (not found):
                        found=True
                        ret_lst.append("female")
                for word in male_prefixes:
                    if (find_exact(row_str, word)) != -1 and (not found):
                            found = True
                            ret_lst.append("male")

                if(not found):
                    # print(row_str)
                    ret_lst.append("unknown")
    df["gender"]=ret_lst
    df.to_excel("genders2.xlsx")

def add_links_to_corpus():
    lst=[]
    urls= open('list_of_artists_urls.txt','r', encoding="utf-8")
    df = pd.read_excel('artists_info_corpus1.xlsx')
    i=0
    for url in urls:
        lst.append(url)
        print(i)
        i=i+1
    urls.close()
    df["urls"]=lst

    df.to_excel("check5.xlsx")



def replace_space(name):
    return name.replace(" ","+")


def change_name_for_wikidata():
    df = pd.read_excel('only_names.xlsx')
    # print(name)
    x=df.apply(lambda row: replace_space(row['Name']),axis=1)
    df['Name']=x
    df.to_excel("named_changed.xlsx")

global i
i=0
def create_wiki_urls_helper(name,driver):
    global i
    i=i+1
    print(i)
    wiki_prefix = "https://www.wikidata.org/wiki/"
    url_to_wiki = "https://www.wikidata.org/w/index.php?search=" + name + "&search=" + name + "&title=Special:Search&go=לדף&ns0=1&ns120=1"
    try:
        driver.get(url_to_wiki)
        id_in_wiki = driver.find_element_by_class_name("wb-itemlink-id")
        id_in_wiki_with_par = id_in_wiki.text
        id_in_wiki = id_in_wiki_with_par[1:len(id_in_wiki_with_par) - 1]
        url_to_wiki2 = wiki_prefix + id_in_wiki
        return url_to_wiki2

    except:
        return url_to_wiki


def create_wiki_urls_col():
    df = pd.read_excel('urls_to_wiki.xlsx')
    option = webdriver.ChromeOptions()
    option.add_argument('headless')
    driver = webdriver.Chrome(path_to_driver, options=option)
    rows = len(df)
    lst=[]
    column=3
    for row in range(rows):
        row_str = str(df.loc[row][column])
        if (row_str == "None"):
            # print("row: "+str(row),"col: "+str(column))
            name=str(df.loc[row][column-1])
            lst.append(create_wiki_urls_helper(name, driver))
        else:
            global i
            i=i+1
            print(i)
            lst.append(row_str)

    df["new_urls"]=lst
    # df['new']=df['new'].apply(lambda cell:create_wiki_urls_helper(cell,driver) )
    driver.quit()
    df.to_excel("urls_to_wiki3.xlsx")



# add_links_to_corpus()

create_wiki_urls_col()


