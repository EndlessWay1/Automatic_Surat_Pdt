# made by nickson
import bs4
import json
import requests
import re

urls = 'https://docs.google.com/document/d/1MQTO145JhZrd0PQbmIiZAvuTMj-D9mK03_nsUrHHqu0/edit?usp=sharing'

# Allowed char
regex = r'([\w\.()@!#$%\^\&\*\[\]:\-\+=_{};\'‘’\"“”,?|<>~`]+)'
# Turning Text --> Docs
class TextDocs():
    # Constructor acc urls
    def __init__(self, urls):
        self.url = urls
        if 'docs' not in self.url: self.status_code = 404; return 
        # get the response
        try:
            respond = requests.get(self.url)
            self.status_code = respond.status_code
        except Exception as e:
            raise e            
        
        soup = bs4.BeautifulSoup(respond.text, "html.parser")
        self.response = respond
        self.soup = soup
        
        def get_docs_text():
            
            script_with_nonce = self.soup.find_all('script', nonce=True)

            json_file = []
            for scrpt in script_with_nonce:
                # Parsing iff there is this key word
                if "DOCS_modelChunk = " in scrpt.text:
                    docs_script = scrpt.text.split("; DOCS_modelChunkLoadStart",1)[0]
                    
                    json_file.append(docs_script.removeprefix("DOCS_modelChunk = "))
                    
            if not json_file: raise FileNotFoundError("HTML Doesnt Exist")

            # JSON Key = 'chunk','revision'
            return [json.loads(i) for i in json_file]

        self.docs_json = get_docs_text()
        self.docs_text = ''
        # joining into text
        # There is a chance of multiple 'chunk' list
        for d_json in self.docs_json:
            # iterating to the content
            for i in d_json['chunk']:
                if 's' in i:
                    if self.docs_text:
                        if i['s']:
                            self.docs_text += "\n" + i['s']
                            continue
                    self.docs_text = i['s']
        if not self.docs_text: raise ValueError("No Text Detected")


        doc_tex = []
        # Joining into a list
        self.docs_splits = self.docs_text.split('\n') 
        for i in self.docs_splits:
            res = re.findall(regex, i)
            if not res:
                continue
            # print(res)
            doc_tex.append(" ".join(res))

        self.doc_list = doc_tex        

if __name__ == "__main__":
    docs_1 = TextDocs(urls)
    print(docs_1.doc_list)
    docs_2 = TextDocs("https://docs.google.com/document/d/11_szsDV4UEYY6dUI7Bv_-k59Dudgyq5M4KXM6wMgyfo/edit?usp=sharing")
    print(docs_2.doc_list)
    