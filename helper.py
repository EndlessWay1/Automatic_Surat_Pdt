# made by nickson
from parser import TextDocs
import re
import json
from collections import defaultdict

liturgi_url = "https://docs.google.com/document/d/1MQTO145JhZrd0PQbmIiZAvuTMj-D9mK03_nsUrHHqu0/edit?usp=sharing"
Pdt_url = 'https://docs.google.com/document/d/11_szsDV4UEYY6dUI7Bv_-k59Dudgyq5M4KXM6wMgyfo/edit?usp=sharing'

class Liturgi():
    def __init__(self, url):
        try:
            self.liturgi = TextDocs(url)
        except Exception as e:
            print(f"Exception {e} has been raised")
            raise e
        
        self._MONTHS_TO_INT = {
            "Januari": "1",
            "Februari": "2",
            "Maret": "3",
            "April": "4",
            "Mei": "5",
            "Juni": "6",
            "Juli": "7",
            "Agustus": "8",
            "September": "9",
             "Oktober": "10",
             "November": "11",
             "Desember": "12" 
        }
        
    def get_month(self):

        set_of_keys = set(['Tema', 'Ayat Firman', 'Ayat KP', 'Ayat BA', 'Ayat Persembahan'])
        docs_list = self.liturgi.doc_list
        length = len(docs_list)
        
        
        j = 0
        set_of_mon = defaultdict(dict)
        while(j < length):
            if (day_mon_yer := re.match(r'^(\d+) (\w+) (\d+).*?', docs_list[j])):
                semi_dict = {}
                # print(docs_list[j])
                j += 1

                while j < length and not re.match(r'^(\d+) (\w+) (\d+)?', docs_list[j]):
                    if (keys_in_set := re.match(r'^([A-Za-z ]+):(.*?)$', docs_list[j])):
                        # print(keys_in_set.group(1))
                        grp1 = keys_in_set.group(1).strip()
                        if grp1 in set_of_keys:
                            new_text = re.sub(r'[\u201c\u201d]', '"', keys_in_set.group(2))
                            new_text = re.sub(r'[\u2019\u2018]', "'", new_text)
                            semi_dict[grp1] = new_text.strip()

                        # utk naro fokus ke list
                        elif grp1 == 'Fokus':
                            curr_fok = []
                            j += 1
                            while(j < length and (matches := re.match(r'^([^\:0-9]+)$', docs_list[j]))):
                                new_text = re.sub(r'[\u201c\u201d]', '"', matches.group(1))
                                new_text = re.sub(r'[\u2019\u2018]', "'", new_text)
                                curr_fok.append(new_text.strip())
                                j += 1
                            if curr_fok:
                                j -= 1
                            semi_dict[grp1] = curr_fok

                        elif grp1 == 'Lagu':
                            song_dic = {}
                            while j < length and not re.match(r'^(\d+) (\w+) (\d+).*?', docs_list[j]):
                                j += 1
                                while j < length and (songs := re.match(r'(\d+)\.(.*?)$', docs_list[j])):    
                                    new_text = re.sub(r'[\u201c\u201d]', '"', songs.group(2))
                                    new_text = re.sub(r'[\u2019\u2018]', "'", new_text)
                                    song_dic[songs.group(1)] = new_text.strip()
                                    j += 1
                                    
                            semi_dict[grp1] = song_dic
                            continue
                    j += 1
                set_of_mon[self._MONTHS_TO_INT[day_mon_yer.group(2).title()]][day_mon_yer.group(1)] = semi_dict
            else:
                # print(docs_list[j])
                j+=1

        self.Dict_Months = set_of_mon


class Jadwal():
    def __init__(self, url):
        try:
            self.liturgi = TextDocs(url)
        except Exception as e:
            print(f"Exception {e} has been raised")
            return
        self._MONTHS_SET = set([
        "JANUARI",
        "FEBRUARI",
        "MARET",
        "APRIL",
        "MEI",
        "JUNI",
        "JULI",
        "AGUSTUS",
        "SEPTEMBER",
        "OKTOBER",
        "NOVEMBER",
        "DESEMBER"])

        self._MONTHS_TO_INT = {
        "JANUARI": "1",
        "FEBRUARI": "2",
        "MARET": "3",
        "APRIL": "4",
        "MEI": "5",
        "JUNI": "6",
        "JULI": "7",
        "AGUSTUS": "8",
        "SEPTEMBER": "9",
        "OKTOBER": "10",
        "NOVEMBER": "11",
        "DESEMBER": "12" 
        }
    
    def get_jadwal(self):
        docs_list = self.liturgi.doc_list

        jadwal = {}
        length = len(docs_list)
        # print(docs_list)
        i = 0
        while i < length:
            if mathcs := re.match(r'^([^0-9 ]+)',docs_list[i]):
                if mathcs.group(1) in self._MONTHS_SET:
                    mon_dic = {} 
                    curr_mon = mathcs.group(1)
                    i += 1
                    while i < length and docs_list[i] in ['Tanggal', 'Pendeta']:
                        i += 1
                    # print(docs_list[i])
                    while i < length and (number := re.match(r'^(\d+)$', docs_list[i])):
                        i += 1
                        s = []
                        while i < length and not re.match(r'^([^0-9 ]+ \d+)', docs_list[i]) and not re.match(r'^(Tanggal|Pendeta)$', docs_list[i]) and not re.match(r'^(\d+)$', docs_list[i]):
                            s.append(docs_list[i])
                            i += 1
                        mon_dic[number.group(1)] = " ".join(s)
                    
                    
                    jadwal[self._MONTHS_TO_INT[curr_mon]]= mon_dic
                else:
                    i += 1
            else:
                i +=1
        self.Dict_Months = jadwal
                
                        


        

        
if __name__ == "__main__":
    '''
    Liturgi
    Usage = Liturgi (Link)
    Usage.get_month() 
    Usage.Litutgi.doc_list --> raw data
    Usage.Dict_Month --> Key char 1-12 indicating month, inside a key with the string indicating the days

    inside the day --> Usage.Dict_Month[month][days]['Tema', 'Ayat Firman', 'Ayat KP', 'Ayat BA', 'Ayat Persembahan']: string with values indicating to it 
    Usage.Dict_Month[month][days]['Fokus']: List, 
    Usage.Dict_Month[month][days]['Tema']: dict--> 1-6 or 1-7 dependending

    Jadwal
    Usage = Jadwal(Link)
    Usage.get_month()
    Usage.Jadwal.doc_list --> raw data
    Usage.Dict__Month --> Key char 1-12 indicating month, inside a key with the string indicating the days
    Usage.Dict_Month[month][days] --> PF pada hari itu
    '''
     
    liturgi = Liturgi(url=liturgi_url)

    # print(liturgi.liturgi.doc_list)
    liturgi.get_month()
    jadwal = Jadwal(Pdt_url)
    jadwal.get_jadwal()

    # print(json.dumps(liturgi.Dict_Months['11']['30'], indent=4))
    # print(json.dumps(liturgi.Dict_Months['7'], indent=4))
    # print(json.dumps(liturgi.Dict_Months['3'], indent=4))
    # print(json.dumps(liturgi.Dict_Months['8'], indent=4))
    print(json.dumps(liturgi.Dict_Months, indent=4))
    print(json.dumps(jadwal.Dict_Months, indent=4))

    
