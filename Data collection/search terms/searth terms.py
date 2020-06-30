import os
import pandas as pd

def extract_url(path):
    df = pd.read_excel(path, usecols=[1], names=None)  # 读取项目名称列,不要列名
    df_li = df.values.tolist()
    result = []
    for s_li in df_li:
        result.append(s_li[0])
    urls = []
    for url in result:
        #strip the final '/'
        if url[-1] == '/':
            url1 = url[:-1]
        else:
            url1 = url

        #delete 'http' and 'www'
        if url1[:7] == 'http://':
            url2 = url1[7:]
        else:
            url2 = url1

        if url2[:8] == 'https://':
            url3 = url2[8:]
        else:
            url3 = url2

        if url3[:4] == 'www.':
            url4 = url3[4:]
        else:
            url4 = url3

        #print(url4)
        urls.append('"'+ url4 + '"')
    print(len(urls))
    return urls

def url_to_searchtearm(url_list,path):
    print(url_list)
    base_term1 = 'ALL('
    base_term2 = ') AND PUBYEAR > 2017 AND ( LIMIT-TO ( DOCTYPE , "ar" ) OR LIMIT-TO ( DOCTYPE , "cp" ) ) AND ( LIMIT-TO ( LANGUAGE , "English" ) )'
    terms = ''
    for url in url_list:
        terms = terms + url + ' OR '
    search_term = base_term1 + terms[:-4] + base_term2
    print(search_term)
    f = open(path,'w')
    f.write(search_term)
    f.close()
    return



if __name__ == '__main__':
    path_OGD = 'OGD portols.xlsx'
    path_search_term = 'scopus.txt'
    url_list = extract_url(path_OGD)
    url_to_searchtearm(url_list,path_search_term)


