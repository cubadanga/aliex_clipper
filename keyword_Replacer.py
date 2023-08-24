import pandas as pd
import sys
import time
from colorama import init, Fore

def read_excel():

    try:
        df_all = pd.read_excel('./학습단어매크로.xlsm', sheet_name = '누적관리', header = 0)
        df_all = df_all.fillna('')
        
    except ValueError as e:
        print(Fore.RED + '오류 - 엑셀 시트의 시트명이 다르거나 올바른 파일이 아닙니다.'+'\n')
        print(Fore.RESET + "엔터를 누르면 종료합니다.")
        aInput = input("")
        sys.exit()

    except FileNotFoundError as e:
        print(Fore.RED + '수정_학습단어매크로.xlsm 파일을 찾을 수 없습니다.'+'\n'+'이런 경우, 파일명이 잘못된 경우가 대부분이었습니다.'+' 이 파일은 필수 파일입니다.'+'\n')
        print(Fore.RESET + "엔터를 누르면 종료합니다.")
        aInput = input("")
        sys.exit()
    return df_all

def edit_excel(df_all):
    df = df_all[['source','keyword','type']]
    is_ban = df[df['type'] == '금지 단어']
    #source_df = is_ban['Source']
    #type_df = is_ban['type']
    is_ban_keyword = is_ban['keyword']
    ban_df = pd.DataFrame({'사이트':['공용']*len(is_ban_keyword),'타입':['금지 단어']*len(is_ban_keyword),'값':is_ban_keyword})

    del_word = df[df['type'] == '제외 단어']
    #source_df = del_word['Source']
    #type_df = del_word['type']
    del_keyword = del_word['keyword']
    del_df = pd.DataFrame({'사이트':['공용']*len(del_keyword),'타입':['제외 단어']*len(del_keyword),'값':del_keyword})

    exchange_word = df[(df['type'] != '제외 단어') & (df['type'] != '금지 단어')]
    #source_df = exchange_word['Source']
    ex_keyword = exchange_word['keyword']
    type_df = exchange_word['type']
    ex_df  = pd.DataFrame({'사이트':['공용']*len(ex_keyword),'타입':['학습 단어']*len(ex_keyword),'값1':ex_keyword, '값2':type_df})

    return ban_df, del_df, ex_df

def save_vidos():
    ban_df.to_excel('./금지단어'+'_'+ tday_s +'.xlsx',index=False)
    del_df.to_excel('./제외단어'+'_'+ tday_s +'.xlsx',index=False)
    ex_df.to_excel('./학습단어'+'_'+ tday_s +'.xlsx',index=False)
    print('\n'+ Fore.LIGHTBLUE_EX + "비도스용 완성! 엔터를 누르면 종료합니다." + Fore.RESET)
    aInput = input("")

def save_coudae(ban_df, del_df, ex_df):
    ban_df = ban_df.rename(columns={'값':'값1'})
    del_df = del_df.rename(columns={'값':'값1'})
    concat_df = pd.concat([ban_df,del_df,ex_df], ignore_index=True)
    type_df = concat_df['타입']
    site_df = concat_df['사이트']
    value1 = concat_df['값1']
    value2 = concat_df['값2']

    coudae_df = pd.DataFrame({'타입':type_df,'사이트':site_df,'값1':value1, '값2':value2})
    coudae_df.to_excel('./쿠대금지단어'+'_'+tday_s+'.xlsx',index=False)
    print('\n'+ Fore.LIGHTBLUE_EX + "쿠대숑숑용 완성! 엔터를 누르면 종료합니다." + Fore.RESET)
    aInput = input("")

tday = time.time()
tday_s = time.strftime('%Y%m%d%H%M',time.localtime(time.time()))
df_all = read_excel()
ban_df, del_df, ex_df = edit_excel(df_all)
save_vidos()

# 프로그램 별 선택
# print(Fore.LIGHTBLUE_EX + '비도스용은 숫자 0 입력, 쿠대숑숑은 숫자 1 입력'+'\n'+ Fore.RESET)
# aInput = input()

# if aInput == '0':
#     save_vidos()

# elif aInput == '1':
#     save_coudae(ban_df, del_df, ex_df)

# else:
#     print("숫자를 똑바로 입력하세요.")



