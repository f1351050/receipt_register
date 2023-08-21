from PIL import Image
import sys
import pyocr
import pyocr.builders
import pandas as pd
import re
from tkinter import filedialog

class ocr:
    def __init__(self,path,df,category,tempo):
        self.path=path
        self.df=df
        self.category_df=category
        self.tempo_df=tempo
    def main(self):
        txt=self.ocr(self.path)
        ##----------文字列から配列へ----------------------------##
        self.moziretu=txt.split("\n")
        ##----------初期値-------------------------------------##
        self.kirokuNo=self.df.iloc[-1]['kirokuNo']
        self.tempo='不明'
        self.hizuke='不明'
        cols = ['kirokuNo','tempo', 'hizuke','dai_bunrui','tyuu_bunrui','no','syouhinmei','kazu','tani','net_value','memo']
        self.result_df = pd.DataFrame(index=[], columns=cols)
        
        try:
            self.hizuke=self.hizuke_extract()
            self.tempo=self.tempo_extract()
        except Exception as e:
            pass
        # try:
        self.date_extract()
        self.result_df=self.result_df.apply(self.category_bunrui,cate_df = self.category_df,axis=1)
        # except Exception as e:
            # pass    

        return txt,self.moziretu,self.result_df
    
    ##---------OCR--------------------------------##
    def ocr(self,img):
        tools=pyocr.get_available_tools()
        if len(tools)==0:
            print("No OCR tool found")
            sys.exit(1)
        tool=tools[0]

        txt=tool.image_to_string(
            Image.open(img),
            lang="jpn",
            builder=pyocr.builders.TextBuilder(tesseract_layout=6)
            )
        return txt
    ##---------日付抽出----moziretu➡hizuke--------------------------------##
    def hizuke_extract(self):   
        _hizuke = [s for s in self.moziretu if re.match(r'\d{4}年\d{1,2}月\d{1,2}日',s)][0].split()
        hizuke=[s for s in _hizuke if re.match(r'\d{4}年\d{1,2}月\d{1,2}日',s)][0]
        none_date=re.sub('年|月','/',hizuke)
        if none_date.find('日')!=-1:
            none_date=none_date[:none_date.rfind('日')]
            hizuke=none_date
        return hizuke
    ##---------店舗の抽出---------------------------------------------------##
    def tempo_extract(self):          
        for index,date in self.tempo_df.iterrows():
            if pd.isnull(date["kouho"])==False:
                s=[s for s in self.moziretu if s.find(date["kouho"]) !=-1]
                if not len(s)==0:
                    tempo=date["tempo_mei"]
        return tempo
    ##----------商品名・金額・数量(kazu)の抽出-------------------------------------##
    def tuika(self,kirokuNo,syouhinmei,kazu,kingaku,memo):
        result_temp = pd.DataFrame({"kirokuNo":[kirokuNo],
                                    "tempo":[self.tempo],
                                    "hizuke":[self.hizuke],
                                    "dai_bunrui":['不明'],
                                    "tyuu_bunrui":['不明'],
                                    "no":[''],
                                    "syouhinmei":[syouhinmei],
                                    "kazu":[kazu],
                                    "tani":[''],
                                    "net_value":[kingaku],
                                    "memo":[memo]})
        self.result_df = pd.concat([self.result_df,result_temp])

    def date_extract(self):
        moziretu=self.moziretu
        first_index=moziretu.index([s for s in moziretu if (s.find('\\') !=-1) or (s.find('殺') !=-1)][0])
        next_mozi=moziretu[first_index][moziretu[first_index].rfind('\\')+1]
        while next_mozi=='' or next_mozi.isdigit()==False:
            moziretu=moziretu[first_index+1:]
            first_index=moziretu.index([s for s in moziretu if (s.find('\\') !=-1)][0])
            next_mozi=moziretu[first_index][moziretu[first_index].rfind('\\')+1]
        moziretu=moziretu[first_index:]
        final_index=moziretu.index([s for s in moziretu if (s.find('計')!=-1) or (s.find('消費')!=-1)][0])
        moziretu=moziretu[:final_index]
        # print(moziretu)
        for i,sk in enumerate(moziretu):
            self.kirokuNo +=1
            memo=''
            kazu=1
            t =1
            kingaku=0
            if sk!='' and sk.find('単')==-1 and sk.find('値引')==-1 and sk.find('%')==-1:
                if i<len(moziretu)-1:
                    while moziretu[i+t]=='':
                        t +=1
                    plus_index=moziretu[i+t]
                    if plus_index.find('値引')!=-1 or plus_index.find('%')!=-1:
                        index=plus_index.rfind('%')
                        memo=plus_index[index-2:index+1]
                    elif plus_index.find('単')!=-1:
                        continue
                sk= sk.split()
                if len(sk)<2:
                    continue
                if re.search(r'\d',sk[-1])!=None:
                    kingaku=re.sub(r'\D','',sk[-1])
                syouhinmei=sk[-2]
                if len(sk)>2:
                    if re.search(r'([ぁ-ん]+|[ァ-ヴー]+|[一-龠]+){2}',sk[-3]):
                        syouhinmei =sk[-3] + syouhinmei
                if re.search('@',syouhinmei[0]):
                        syouhinmei=syouhinmei[1:]
            elif sk.find('単')!=-1 or sk.find('コ')!=-1:
                while moziretu[i-t]=='' and i-t>0:
                    t +=1
                if i-t>0:
                    mainus_index=moziretu[i-t]
                else:
                    i=self.moziretu.index(sk)
                    while self.moziretu[i-t]=='':
                        t -=1
                    mainus_index=self.moziretu[i-t]
                k_syouhinmei=mainus_index.split()
                syouhinmei=k_syouhinmei[-1]
                if len(k_syouhinmei)>1:
                    if re.search(r'([ぁ-ん]+|[ァ-ヴー]+|[一-龠]+){2}',k_syouhinmei[-2]):
                            syouhinmei =sk[-2] + syouhinmei

                tk=sk.split()
                kingaku=re.sub(r'\D','',tk[-1])
                kazu=tk[-2]
                kazu=kazu[:kazu.rfind('コ')]
            else:
                continue
            self.tuika(self.kirokuNo,syouhinmei,kazu,kingaku,memo)

        
    ##----------単位、中分類、大分類をcategoryから抽出-------------------------------------##
    def category_bunrui(self,r_df,cate_df):
        temp_cate = cate_df[cate_df["syokuhin_mei"].str.contains(r_df["syouhinmei"],case=False,regex=False)== True]
        if len(temp_cate) != 0:
            r_df["tyuu_bunrui"] = temp_cate.iloc[0]["tyuu_bunrui"]
            r_df["dai_bunrui"] = temp_cate.iloc[0]["dai_bunrui"]
            r_df["no"] = temp_cate.iloc[0]["No"]
            r_df["tani"] = temp_cate.iloc[0]["tani"]
        return r_df

def file_path_get():
        filename = filedialog.askopenfilename(
            filetypes=[("Image file", ".bmp .png .jpg .tif"), ("Bitmap", ".bmp"), ("PNG", ".png"), ("JPEG", ".jpg"), ("Tiff", ".tif") ],
            initialdir=r"C:\Users\f1351\Desktop\レシート"
            )
        return filename
def rireki_excel_get(filepath):
    try:
        df=pd.read_excel(filepath,sheet_name='rireki',index_col=None)
        category_df=pd.read_excel(filepath,sheet_name='category',index_col=None)
        tempo_df=pd.read_excel(filepath,sheet_name='tempo',index_col=None)           
    except Exception as e:
        print(e)
    return df,category_df,tempo_df


if __name__ == '__main__':
    filepath=r"C:\Users\f1351\Desktop\rireki.xlsx"
    imag_path=file_path_get()
    img = Image.open(imag_path)
    img.show()
    print(imag_path)

    df,category_df,tempo_df=rireki_excel_get(filepath)
    # imag_path=r'C:\Users\f1351\AppData\Local\Temp\tmpniao37nc.PNG'
    root=ocr(imag_path,df,category_df,tempo_df)

    txt,moziretu,result_df=root.main()
    if len(result_df)>0:
        print(result_df)
    else:
        print('エラーのため、抽出できませんでした。')
    print('---------')
    print(imag_path)
    print('---------')
    print(moziretu)
    # shutil.move(imag_path,r"C:\Users\f1351\Desktop\レシート\スキャン確認済み\test")