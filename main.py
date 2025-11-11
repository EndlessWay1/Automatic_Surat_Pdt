# made by nickson
from helper import Liturgi, Jadwal
import doc_maker as dm
import sys
import datetime
import os
import json
import re

liturgi_url = "https://docs.google.com/document/d/1MQTO145JhZrd0PQbmIiZAvuTMj-D9mK03_nsUrHHqu0/edit?usp=sharing"
Pdt_url = 'https://docs.google.com/document/d/11_szsDV4UEYY6dUI7Bv_-k59Dudgyq5M4KXM6wMgyfo/edit?usp=sharing'

def main():
    try:
        liturgi = Liturgi(liturgi_url)
        liturgi.get_month()
    except Exception as e:
        print(f"Exception Error {e} has occured")
        sys.exit()
    
    try:
        jadwal = Jadwal(Pdt_url)
        jadwal.get_jadwal()
    except Exception as e:
        print(f"Exception Error {e} has occured")
        sys.exit()
    
    # print(json.dumps(liturgi.Dict_Months['12'], indent = 4))
    # print(liturgi.liturgi.doc_list)
    try:
        Months = int(input("Months: "))
    except ValueError:
        print("Months must be an int")
        sys.exit()

    try:
        year_due = int(input("Year: "))
    except ValueError:
        print("Year must be an int")
        sys.exit()
    
    try:
        No_surat = int(input("Nomor Surat: "))
    except ValueError:
        print("Nomor Surat must be an int")
        sys.exit()
        
    file_zip_loc(liturgi, jadwal, No_sur=No_surat, year_due=year_due, month=Months)


def file_zip_loc(liturgi: Liturgi, jadwal: Jadwal, No_sur, path= os.path.dirname(os.path.realpath(__file__)),year_due = str(datetime.date.today().year), month = str(datetime.date.today().month + 1 if datetime.date.today().month + 1 < 13 else 1)):
    if not (liturgi and jadwal): 
        print("Incorrect Usage of Function, Liturgi and Jadwal needs to be filled with corrects Parms")
        return
    
    month = str(month)

    year_due = str(year_due)
    pdt_dalam = {'Naya':os.path.join(os.path.dirname(os.path.realpath(__file__)), "static/Surat Pendeta Pak_Naya.docx"), 'Gloria':os.path.join(os.path.dirname(os.path.realpath(__file__)), "static/Surat Pendeta Kak_Gloria.docx"), 'Angela':os.path.join(os.path.dirname(os.path.realpath(__file__)), "static/Surat Pendeta Kak_Angela.docx")}
    tanggal = jadwal.Dict_Months[month].keys()
    curr_lit = liturgi.Dict_Months[month]
    curr_jad = jadwal.Dict_Months[month]

    INT_TO_MONTHS = {
    '1' : "Januari",
    '2' : "Februari",
    '3' : "Maret",
    '4' : "April",
    '5' : "Mei",
    '6' : "Juni",
    '7' : "Juli",
    '8' : "Agustus",
    '9' : "September",
    '10' : "Oktober",
    '11' : "November",
    '12' : "Desember" 
    }

    s_months = INT_TO_MONTHS[month]
    # print(curr_lit)
    print(s_months,':')
    for t in tanggal:
        print(f"{t} {s_months}")
        if t in curr_jad:
            print(f"Pelayan Firman: {curr_jad[t] if curr_jad[t] else 'TIDAK ADA'}")
        else:
            print(f"PF pada tanggal {t} {s_months}, tidak ada.\n")
        if t in curr_lit:
            print(f"Tema: \"{curr_lit[t]['Tema'] if 'Tema' in curr_lit[t] else ''}\"")
            print(f"Ayat Firman: {curr_lit[t]['Ayat Firman'] if 'Ayat Firman' in curr_lit[t] else ''}")
            print("Fokus:")
            if 'Fokus'in curr_lit[t]:
                for Foc in curr_lit[t]['Fokus']:
                    print(f"●   {Foc}")
            print()
        else:
            print(f"Liturgi pada tanggal {t} {s_months}, tidak ada.\n")
    print(f"Make zip file to: {path}")
    ans = input("Confirm (Y/N): ").lower().strip()
    if not(ans == "yes" or ans == 'y'): return

    dir_path = dm.make_dir(path=path, s_mon=s_months,years=year_due)
    # print(dir_path)

    if dir_path == -1: return
    elif dir_path == 1:
        cur_path = os.path.join(path,f'{s_months} {year_due}')
        print(f"Folder {cur_path} has already been created")
        ans = input("Are You Sure (Y/N): ").lower().strip()
        if not(ans == "yes" or ans == 'y'): return
        
        try:
            dm.delete_dir(cur_path)
        except ValueError:
            print("Unsuccessfull Deleting the file, because there is a directory inside.")
            return
        dir_path = dm.make_dir(path=path, s_mon=s_months,years=year_due)
    
    for t in tanggal:
        if t in curr_jad and t in curr_lit and not ("Tema" not in curr_lit[t] or "Fokus" not in curr_lit[t] or "Ayat Firman" not in curr_lit[t] or not curr_jad[t]):
            if matches := re.search(r'(Angela|Naya|Gloria)', curr_jad[t]):
                Surat_dalam = dm.Surat(path = pdt_dalam[matches.group(1)], directory_name=dir_path, NoSur=No_sur, due=f"{t} {s_months} {year_due}", theme=curr_lit[t]["Tema"].strip(), temaItalic=True, alkitab=curr_lit[t]["Ayat Firman"].strip(), Fokus=curr_lit[t]["Fokus"], NamPF=curr_jad[t], years=year_due, mon=month)
                Surat_dalam.make_document()
                Surat_dalam.add_text()
                Surat_dalam.save_document()
                print(f"Surat Dalam {curr_jad[t]} pada tanggal {t} {s_months} {year_due}, telah dibuat.")

            else:
                Surat_luar = dm.Surat(path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "static/Surat Pendeta Luar.docx"), directory_name=dir_path, NoSur=No_sur, due=f"{t} {s_months} {year_due}", theme=curr_lit[t]["Tema"].strip(), temaItalic=True, alkitab=curr_lit[t]["Ayat Firman"].strip(), Fokus=curr_lit[t]["Fokus"], NamPF=curr_jad[t], years=year_due, mon=month)
                Surat_luar.make_document()
                Surat_luar.add_text()
                Surat_luar.save_document()
                print(f"Surat Luar {curr_jad[t]} pada tanggal {t} {s_months} {year_due}, telah dibuat.")
            No_sur += 1
        else:
            print(f"Surat pada tanggal {t} {s_months} {year_due}, telah dibatalkan karena PF atau Liturgi yang tidak lengkap")
        
    dm.zipping(dir_path, dir_path)
    dm.delete_dir(dir_path)
    print("Success!!!")
    return 0

    

if __name__ == "__main__":
    main()