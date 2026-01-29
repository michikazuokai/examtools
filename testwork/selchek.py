import pandas as pd
from pathlib import Path
import random
import string

def shuffle(arr):
    indexed_arr = [(element, index) for index, element in enumerate(arr)]
    random.shuffle(indexed_arr)
    mapping = {}
    for new_idx, (element, orig_idx) in enumerate(indexed_arr):
        mapping[orig_idx] = new_idx
    return [v[0] for v in indexed_arr],mapping

# ラムダ関数で簡潔に
alp2n = lambda c: ord(c) - ord('A')

def shuffle_choices(df):
    anscnt=0
    sanscnt=0
    i = 0
    n = len(df)
    selst=[]
    selflg=False
    while i < n:
        code = df.iloc[i, 0]

        if "【" in code:
            svcode=code

        if code == "b_select":
            print(svcode)
            selflg=True
            print('---  select ---')

        if selflg and code == "select":
            selst.append(df.iloc[i,1])

        if selflg and code == "e_select":
            print(selst)
            narr,nmap =shuffle(selst)
            print(narr)
            print(nmap)

        if selflg and code == "answer":
            al=(df.iloc[i,1]).split(',')
            for v in al:
                print(v, alp2n(v),nmap[alp2n(v)],chr(nmap[alp2n(v)]+65))
            anscnt+=1

        if selflg and code == "e_answer":
            print('@@@ ans end ---')
            selflg=False
            anscnt=0
            selst=[]

        if code == "b_subselect":
            print(svcode)
            print('---  sub select ---')
            selflg=True

        if selflg and code == "subselect":
            selst.append(df.iloc[i,1])

        if selflg and code == "e_subselect":
            print(selst)
            narr,nmap =shuffle(selst)
            print(narr)
            print(nmap)

        if selflg and code == "subanswer":
            al=(df.iloc[i,1]).split(',')
            for v in al:
                print(al, alp2n(v),nmap[alp2n(v)],chr(nmap[alp2n(v)]+65))
            anscnt+=1

        if selflg and code == "e_subanswer":
            print("@@@ sub answer end ---")
            selflg=False
            selst=[]
            sanscnt=0

        i += 1


# ====== 実行例 ======
# 現在の場所
curdir = Path(__file__).parent
df = pd.read_excel(curdir / "exam_1.xlsx")
shuffle_choices(df)