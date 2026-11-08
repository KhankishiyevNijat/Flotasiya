import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.comments import Comment

F="Arial"
NAVY="1F3864"; SEC="D9E1F2"; INP="FFF2CC"; OUT="E2EFDA"; GREY="F2F2F2"
thin=Side(style="thin",color="BFBFBF")
BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
def font(**k): k.setdefault("name",F); k.setdefault("size",10); return Font(**k)
def fill(c): return PatternFill("solid",fgColor=c)
C=Alignment(horizontal="center",vertical="center",wrap_text=True)
L=Alignment(horizontal="left",vertical="center",wrap_text=True)

def title(ws,ref,text,sub=None,subref=None):
    ws.merge_cells(ref); c=ws[ref.split(":")[0]]
    c.value=text; c.font=font(size=16,bold=True,color="FFFFFF"); c.fill=fill(NAVY); c.alignment=L
    ws.row_dimensions[c.row].height=30
    if sub:
        ws.merge_cells(subref); s=ws[subref.split(":")[0]]
        s.value=sub; s.font=font(size=9,italic=True,color="595959"); s.alignment=L
def section(ws,row,text,c1=1,c2=4):
    ws.merge_cells(start_row=row,start_column=c1,end_row=row,end_column=c2)
    c=ws.cell(row,c1,text); c.font=font(size=11,bold=True,color=NAVY); c.fill=fill(SEC); c.alignment=L
    ws.row_dimensions[row].height=20
def inp(c,fmt=None):
    c.font=font(color="0000FF",bold=True); c.fill=fill(INP); c.border=BOX; c.alignment=C
    if fmt: c.number_format=fmt
def out(c,fmt=None,bold=False):
    c.font=font(bold=bold); c.fill=fill(OUT); c.border=BOX; c.alignment=C
    if fmt: c.number_format=fmt

wb=Workbook()

# ================= KALKULYATOR =================
k=wb.active; k.title="Kalkulyator"; k.sheet_view.showGridLines=False
for col,w in zip("ABCDE",[42,16,10,62,2]): k.column_dimensions[col].width=w
title(k,"A1:D1","  Tail hədəf kalkulyatoru — Cu flotasiya (AIMC Agdere)",
      "Sarı xanalar daxil edilir (göy rəqəm). Yaşıl xanalar avtomatik hesablanır. Nümunə: 8 noyabr 23:00 analizi.","A2:D2")
def hdr(ws,row,cols):
    for i,t in enumerate(cols,1):
        c=ws.cell(row,i,t); c.font=font(bold=True,color="FFFFFF"); c.fill=fill("44546A"); c.alignment=C; c.border=BOX

section(k,4,"1. GİRİŞ MƏLUMATLARI")
hdr(k,5,["Parametr","Dəyər","Vahid","Qeyd"])
inputs=[
 (6,"Nümunə tarixi / saatı","8 noyabr 23:00","","Mətn"),
 (7,"Tonaj",450,"t/saat","Hazırkı emal sürəti"),
 (8,"Flotation feed Cu",0.646,"%","Lab analizi"),
 (9,"Flotation feed CuOx",0.1019,"%","Oksid mis analizi"),
 (10,"Total tail Cu",0.293,"%","Rougher tail (total) — cleaner scavenger daxil"),
 (11,"Tail CuOx",0.074,"%","Oksid mis analizi"),
 (12,"Konsentrat Cu",18.0,"%","Son conc / thickener U/F. 8 noyabr üçün fərzi dəyər"),
 (13,"İstənilən tail Cu (yoxlama üçün)",0.10,"%","Reallıq yoxlaması: bu tail-ə çatmaq olarmı?"),
]
for r,lab,v,u,n in inputs:
    k.cell(r,1,lab).font=font(); k.cell(r,1).border=BOX; k.cell(r,1).alignment=L
    c=k.cell(r,2,v); inp(c,"0.000" if isinstance(v,float) and v<1 else ("0.00" if isinstance(v,float) else "0"))
    k.cell(r,3,u).font=font(color="595959"); k.cell(r,3).alignment=C; k.cell(r,3).border=BOX
    k.cell(r,4,n).font=font(size=9,italic=True,color="595959"); k.cell(r,4).alignment=L; k.cell(r,4).border=BOX
k["B9"].number_format="0.0000"; k["B12"].number_format="0.00"

section(k,15,"2. FƏRZİYYƏLƏR (zavod tarixçəsinə görə dəqiqləşdir)")
hdr(k,16,["Parametr","Dəyər","Vahid","Qeyd"])
ass=[
 (17,"Sulfid Cu bərpası — normal iş",0.85,"%","Hədəfin yuxarı həddi üçün"),
 (18,"Sulfid Cu bərpası — çox yaxşı iş",0.90,"%","Hədəfin aşağı həddi üçün"),
 (19,"Cleaner scavenger əlavəsi — min",0.02,"% Cu","Cleaner scav tail-in total tail-ə verdiyi pay"),
 (20,"Cleaner scavenger əlavəsi — max",0.03,"% Cu",""),
 (21,"Feed oksid payı (CuOx yoxdursa)",0.1577,"%","8 noyabr: 0.1019 / 0.646. Gündəlikdə CuOx boş olanda istifadə olunur"),
 (22,"Oksid Cu bərpası (CuOx yoxdursa)",0.27,"%","8 noyabr analizi əsasında ~27%"),
 (23,"Qeyri-real hədd: sulfid bərpası >",0.95,"%","Bundan yuxarı tələb olunursa — real deyil"),
 (24,"Çətin hədd: sulfid bərpası >",0.90,"%",""),
]
for r,lab,v,u,n in ass:
    k.cell(r,1,lab).font=font(); k.cell(r,1).border=BOX; k.cell(r,1).alignment=L
    c=k.cell(r,2,v); inp(c,"0.0%" if u=="%" else "0.000")
    k.cell(r,3,u).font=font(color="595959"); k.cell(r,3).alignment=C; k.cell(r,3).border=BOX
    k.cell(r,4,n).font=font(size=9,italic=True,color="595959"); k.cell(r,4).alignment=L; k.cell(r,4).border=BOX
k["B21"].number_format="0.00%"

section(k,26,"3. NƏTİCƏLƏR")
hdr(k,27,["Göstərici","Nəticə","Vahid","Düstur / izah"])
res=[
 (28,"Feed oksid Cu payı","=IFERROR(B9/B8,0)","0.0%","%","CuOx / Cu (feed)"),
 (29,"Feed sulfid Cu","=B8-B9","0.000","%","Feed Cu − Feed CuOx"),
 (30,"Tail oksid Cu payı","=IFERROR(B11/B10,0)","0.0%","%","Tail CuOx / Tail Cu"),
 (31,"Tail sulfid Cu","=B10-B11","0.000","%","Tail Cu − Tail CuOx"),
 (32,"Mass pull (conc çıxımı)","=IFERROR((B8-B10)/(B12-B10),0)","0.00%","%","(f − t) / (c − t)"),
 (33,"Tail kütlə payı","=1-B32","0.0%","%","1 − mass pull"),
 (34,"Ümumi Cu bərpası","=IFERROR(B12*(B8-B10)/(B8*(B12-B10)),0)","0.0%","%","c(f − t) / f(c − t)"),
 (35,"Sulfid Cu bərpası","=IFERROR(1-B31*B33/B29,0)","0.0%","%","1 − tail sulfid × tail payı / feed sulfid"),
 (36,"Oksid Cu bərpası","=IFERROR(1-B11*B33/B9,0)","0.0%","%","1 − tail CuOx × tail payı / feed CuOx"),
 (37,"HƏDƏF TAIL — aşağı (çox yaxşı iş)","=B11+B29*(1-B18)/B33+B19","0.000","% Cu","Tail CuOx + sulfid itkisi (90%) + cl. scav min"),
 (38,"HƏDƏF TAIL — yuxarı (normal iş)","=B11+B29*(1-B17)/B33+B20","0.000","% Cu","Tail CuOx + sulfid itkisi (85%) + cl. scav max"),
 (39,"Faktiki total tail","=B10","0.000","% Cu","Giriş xanasından"),
 (40,"VƏZİYYƏT",'=IF(B10<=B37,"ƏLA",IF(B10<=B38,"NORMAL",IF(B10<=B38*1.25,"DİQQƏT","PİS")))',None,"","ƏLA ≤ aşağı hədd · NORMAL ≤ yuxarı hədd · DİQQƏT ≤ yuxarı×1.25 · PİS"),
 (41,"Hədəf tail-də bərpa (yuxarı hədd)","=IFERROR(B12*(B8-B38)/(B8*(B12-B38)),0)","0.0%","%","Normal işdə gözlənilən bərpa"),
 (42,"Bərpa boşluğu","=B41-B34","0.0%","faiz bəndi","Hədəfə çatmaq üçün itirilən bərpa"),
 (43,"Tail-ə gedən Cu (faktiki)","=B7*24*B33*B10/100","0.00","t Cu/gün","Tonaj × 24 × tail payı × tail Cu"),
 (44,"Hədəfdən artıq Cu itkisi","=MAX(0,B7*24*B33*(B10-B38)/100)","0.00","t Cu/gün","Yuxarı hədd ilə müqayisədə"),
]
for r,lab,f_,fmt,u,n in res:
    k.cell(r,1,lab).font=font(bold=lab.isupper() or lab.startswith("HƏDƏF")); k.cell(r,1).border=BOX; k.cell(r,1).alignment=L
    c=k.cell(r,2,f_); out(c,fmt,bold=r in (37,38,40,44))
    k.cell(r,3,u).font=font(color="595959"); k.cell(r,3).alignment=C; k.cell(r,3).border=BOX
    k.cell(r,4,n).font=font(size=9,italic=True,color="595959"); k.cell(r,4).alignment=L; k.cell(r,4).border=BOX
k["B40"].font=font(bold=True,size=12)
k.row_dimensions[40].height=24

section(k,46,"4. REALLIQ YOXLAMASI — istənilən tail-ə çatmaq olarmı?")
hdr(k,47,["Göstərici","Nəticə","Vahid","Düstur / izah"])
real=[
 (48,"İstənilən tail Cu","=B13","0.000","%","Giriş B13"),
 (49,"Lazım olan sulfid bərpası","=IFERROR(1-(B13-B11-B19)*B33/B29,0)","0.0%","%","Oksid itkisi və cl. scav min sabit qəbul edilir"),
 (50,"Bu tail-də ümumi bərpa","=IFERROR(B12*(B8-B13)/(B8*(B12-B13)),0)","0.0%","%","c(f − t) / f(c − t)"),
 (51,"QİYMƏT",'=IF(B13<=B11+B19,"MÜMKÜN DEYİL",IF(B49>B23,"REAL DEYİL",IF(B49>B24,"ÇƏTİN","REAL")))',None,"","Oksid + cl. scav-dan aşağı tail mümkün deyil"),
]
for r,lab,f_,fmt,u,n in real:
    k.cell(r,1,lab).font=font(bold=lab.isupper()); k.cell(r,1).border=BOX; k.cell(r,1).alignment=L
    c=k.cell(r,2,f_); out(c,fmt,bold=True)
    k.cell(r,3,u).font=font(color="595959"); k.cell(r,3).alignment=C; k.cell(r,3).border=BOX
    k.cell(r,4,n).font=font(size=9,italic=True,color="595959"); k.cell(r,4).alignment=L; k.cell(r,4).border=BOX
k["B51"].font=font(bold=True,size=12); k.row_dimensions[51].height=24

def status_cf(ws,rng,first):
    for txt,bg,fg in [("ƏLA","C6EFCE","006100"),("REAL","C6EFCE","006100"),("NORMAL","DDEBF7","1F4E78"),
                      ("DİQQƏT","FFEB9C","9C5700"),("ÇƏTİN","FFEB9C","9C5700"),
                      ("PİS","FFC7CE","9C0006"),("REAL DEYİL","FFC7CE","9C0006"),("MÜMKÜN DEYİL","FFC7CE","9C0006")]:
        ws.conditional_formatting.add(rng,FormulaRule(formula=[f'{first}="{txt}"'],fill=fill(bg),font=Font(name=F,bold=True,color=fg),stopIfTrue=True))
status_cf(k,"B40","$B$40"); status_cf(k,"B51","$B$51")
k.conditional_formatting.add("B10",FormulaRule(formula=["$B$10>$B$38*1.25"],font=Font(name=F,bold=True,color="C00000")))

dv=DataValidation(type="decimal",operator="greaterThanOrEqual",formula1="0",showErrorMessage=True,
                  errorTitle="Yanlış dəyər",error="Müsbət rəqəm daxil edin")
k.add_data_validation(dv)
for r in range(7,14): dv.add(f"B{r}")
for r in range(17,25): dv.add(f"B{r}")
k["B12"].comment=Comment("8 noyabr 23:00 üçün conc analizi yoxdur — 18% fərz edilib (thickener U/F 18.36).","Claude")
k.freeze_panes="A4"
k.page_setup.orientation="portrait"; k.page_setup.fitToWidth=1; k.sheet_properties.pageSetUpPr.fitToPage=True; k.page_setup.fitToHeight=0

# ================= GÜNDƏLİK =================
g=wb.create_sheet("Gündəlik"); g.sheet_view.showGridLines=False
cols=[("Tarix",11,"in"),("Saat",7,"in"),("Tonaj\nt/saat",8,"in"),("Feed\nCu %",8,"in"),("Feed\nCuOx %",8,"in"),
      ("Total tail\nCu %",9,"in"),("Tail\nCuOx %",8,"in"),("Cl. scav\ntail Cu %",9,"in"),("Conc\nCu %",8,"in"),
      ("Conc\nS %",7,"in"),("Conc\nFe %",7,"in"),("Conc\nSi %",7,"in"),("Rougher\npH",8,"in"),
      ("Mass\npull",8,"out"),("Cu\nbərpası",8,"out"),("S/Cu",7,"out"),("Fe/Cu",7,"out"),("CuOx\nmənbəyi",9,"out"),
      ("Hədəf tail\naşağı",9,"out"),("Hədəf tail\nyuxarı",9,"out"),("Vəziyyət",10,"out"),("Cu itkisi\nt/gün",9,"out"),
      ("Qeyd",46,"in"),("Etiket",13,"out")]
ncol=len(cols)
g.merge_cells(start_row=1,start_column=1,end_row=1,end_column=ncol)
c=g.cell(1,1,"  Gündəlik — hər analiz dəsti bir sətir"); c.font=font(size=16,bold=True,color="FFFFFF"); c.fill=fill(NAVY); c.alignment=L
g.row_dimensions[1].height=30
g.merge_cells(start_row=2,start_column=1,end_row=2,end_column=ncol)
c=g.cell(2,1,"Sarı sütunları doldur (A–M, W). Yaşıl sütunlar avtomatikdir. CuOx boş qalanda Kalkulyator!B21–B22 fərziyyələri ilə təxmin edilir. Fərziyyələr Kalkulyator vərəqindən götürülür.")
c.font=font(size=9,italic=True,color="595959"); c.alignment=L
g.merge_cells(start_row=3,start_column=1,end_row=3,end_column=13)
c=g.cell(3,1,"DAXİL EDİLƏN"); c.font=font(bold=True,color="7F6000"); c.fill=fill(INP); c.alignment=C
g.merge_cells(start_row=3,start_column=14,end_row=3,end_column=22)
c=g.cell(3,14,"AVTOMATİK HESABLANIR"); c.font=font(bold=True,color="375623"); c.fill=fill(OUT); c.alignment=C
from openpyxl.utils import get_column_letter as CL
for i,(h,w,t) in enumerate(cols,1):
    c=g.cell(4,i,h); c.font=font(bold=True,color="FFFFFF"); c.fill=fill(NAVY if t=="in" else "375623"); c.alignment=C; c.border=BOX
    g.column_dimensions[CL(i)].width=w
g.row_dimensions[4].height=32
data=[
 (dt.date(2026,11,5),"04:00",500,0.668,None,0.099,None,None,19.12,27.48,23.17,2.41,11.44,"03:07 reagent dəyişikliyindən sonra"),
 (dt.date(2026,11,5),"07:20",500,0.607,None,0.116,None,1.66,20.88,26.86,22.44,2.53,11.40,"Conc yaxşılaşdı, bərpa düşdü; cl. scav tail yüksək"),
 (dt.date(2026,11,5),"16:00",500,0.701,None,0.257,None,0.73,23.65,28.20,21.38,2.30,11.50,"Feed dəyişdi: S 5.42, Fe 5.83 — pirit artdı"),
 (dt.date(2026,11,5),"23:00",450,0.426,None,0.264,None,0.825,16.79,30.71,24.93,2.69,11.37,"Feed 22:00 analizi; həddindən artıq üzdürmə"),
 (dt.date(2026,11,6),"01:20",450,0.439,None,0.099,None,None,13.32,32.67,28.31,1.46,11.36,"Feed 01:00; conc pirit ilə durulaşıb (S/Cu 2.45)"),
 (dt.date(2026,11,8),"23:00",None,0.646,0.1019,0.293,0.074,None,18.0,None,None,None,None,"Oksid analizi; conc 18% fərzi"),
 (dt.date(2026,11,9),"02:00",None,0.710,None,0.288,None,None,17.33,29.88,22.31,2.62,11.66,"3rd cleaner conc; CuOx təxmini"),
]
FIRST=5; LAST=FIRST+300-1
K="Kalkulyator!"
for r in range(FIRST,LAST+1):
    band = fill("FFFBEA") if (r-FIRST)%2==0 else fill("FFFFFF")
    for i in range(1,ncol+1):
        c=g.cell(r,i); c.border=BOX; c.font=font(); c.alignment=C
        t=cols[i-1][2]
        c.fill = band if t=="in" else (fill("F3F9EE") if (r-FIRST)%2==0 else fill("FFFFFF"))
    g.cell(r,1).number_format="dd.mm.yyyy"
    for i in (4,5,6,7,8): g.cell(r,i).number_format="0.000"
    g.cell(r,5).number_format="0.0000"
    for i in (9,10,11,12,13): g.cell(r,i).number_format="0.00"
    g.cell(r,23).alignment=L
    feedox=f'IF(E{r}="",D{r}*{K}$B$21,E{r})'
    mp=f"N{r}"
    tailox=f'IF(G{r}="",{feedox}*(1-{K}$B$22)/(1-{mp}),G{r})'
    ok=f'OR(D{r}="",F{r}="",I{r}="")'
    g.cell(r,14,f'=IF({ok},"",IFERROR((D{r}-F{r})/(I{r}-F{r}),""))').number_format="0.00%"
    g.cell(r,15,f'=IF({ok},"",IFERROR(I{r}*(D{r}-F{r})/(D{r}*(I{r}-F{r})),""))').number_format="0.0%"
    g.cell(r,16,f'=IF(OR(I{r}="",J{r}=""),"",J{r}/I{r})').number_format="0.00"
    g.cell(r,17,f'=IF(OR(I{r}="",K{r}=""),"",K{r}/I{r})').number_format="0.00"
    g.cell(r,18,f'=IF(D{r}="","",IF(AND(E{r}<>"",G{r}<>""),"Ölçülüb","Təxmin"))')
    g.cell(r,19,f'=IF({ok},"",IFERROR({tailox}+(D{r}-{feedox})*(1-{K}$B$18)/(1-{mp})+{K}$B$19,""))').number_format="0.000"
    g.cell(r,20,f'=IF({ok},"",IFERROR({tailox}+(D{r}-{feedox})*(1-{K}$B$17)/(1-{mp})+{K}$B$20,""))').number_format="0.000"
    g.cell(r,21,f'=IF(OR(S{r}="",T{r}=""),"",IF(F{r}<=S{r},"ƏLA",IF(F{r}<=T{r},"NORMAL",IF(F{r}<=T{r}*1.25,"DİQQƏT","PİS"))))')
    g.cell(r,22,f'=IF(OR(C{r}="",{mp}=""),"",C{r}*24*(1-{mp})*F{r}/100)').number_format="0.00"
    g.cell(r,24,f'=IF(A{r}="","",TEXT(A{r},"dd.mm")&" "&B{r})')
    for i in (14,15,16,17,18,19,20,21,22,24): g.cell(r,i).font=font(color="000000")
    g.cell(r,21).font=font(bold=True)
for j,row in enumerate(data):
    r=FIRST+j
    for i,v in enumerate(row,1):
        col = i if i<=13 else 23
        if v is not None:
            g.cell(r,col,v); g.cell(r,col).font=font(color="0000FF")
g.cell(FIRST+5,9).comment=Comment("Conc analizi yoxdur — 18% fərz edilib.","Claude")
g.cell(FIRST+6,9).comment=Comment("3rd cleaner conc. Thickener U/F 18.36.","Claude")
status_cf(g,f"U{FIRST}:U{LAST}",f"$U{FIRST}")
g.conditional_formatting.add(f"P{FIRST}:P{LAST}",FormulaRule(formula=[f'AND(ISNUMBER(P{FIRST}),P{FIRST}>1.3)'],font=Font(name=F,bold=True,color="C00000")))
g.conditional_formatting.add(f"L{FIRST}:L{LAST}",FormulaRule(formula=[f'AND(ISNUMBER(L{FIRST}),L{FIRST}>3)'],font=Font(name=F,bold=True,color="C00000")))
g.conditional_formatting.add(f"H{FIRST}:H{LAST}",FormulaRule(formula=[f'AND(ISNUMBER(H{FIRST}),H{FIRST}>0.4)'],font=Font(name=F,bold=True,color="C00000")))
g.conditional_formatting.add(f"M{FIRST}:M{LAST}",FormulaRule(formula=[f'AND(ISNUMBER(M{FIRST}),M{FIRST}>11.8)'],font=Font(name=F,bold=True,color="C00000")))
dv2=DataValidation(type="decimal",operator="greaterThanOrEqual",formula1="0",showErrorMessage=True,error="Müsbət rəqəm daxil edin")
g.add_data_validation(dv2); dv2.add(f"C{FIRST}:M{LAST}")
dvd=DataValidation(type="date",operator="greaterThan",formula1="40000",showErrorMessage=True,error="Tarix daxil edin (gg.aa.iiii)")
g.add_data_validation(dvd); dvd.add(f"A{FIRST}:A{LAST}")
g.freeze_panes=g.cell(FIRST,3)
g.auto_filter.ref=f"A4:{CL(ncol)}{LAST}"
g.column_dimensions["X"].hidden=True
# note row under header explaining red marks
g.page_setup.orientation="landscape"; g.sheet_properties.pageSetUpPr.fitToPage=True; g.page_setup.fitToHeight=0

# ================= TREND =================
t=wb.create_sheet("Trend"); t.sheet_view.showGridLines=False
t.merge_cells("A1:P1"); c=t["A1"]; c.value="  Trend — total tail Cu hədəf zolağına qarşı və Cu bərpası"
c.font=font(size=16,bold=True,color="FFFFFF"); c.fill=fill(NAVY); c.alignment=L; t.row_dimensions[1].height=30
t.merge_cells("A2:P2"); c=t["A2"]
c.value=f"Qrafiklər Gündəlik vərəqinin {FIRST}–{FIRST+29}-cu sətirlərini (30 analiz) göstərir. Daha çox üçün: qrafikə sağ klik → Select Data → aralığı genişləndir."
c.font=font(size=9,italic=True,color="595959")
N=FIRST+29
cats=Reference(g,min_col=24,min_row=FIRST,max_row=N)
lc=LineChart(); lc.title="Total tail Cu vs hədəf"; lc.y_axis.title="Cu, %"; lc.x_axis.title=None
lc.height=9; lc.width=24
for col,color,dash,w in [(6,"C00000",None,28000),(19,"70AD47","dash",19050),(20,"ED7D31","dash",19050)]:
    lc.add_data(Reference(g,min_col=col,min_row=4,max_row=N),titles_from_data=True)
    s=lc.series[-1]; s.graphicalProperties.line.solidFill=color; s.graphicalProperties.line.width=w
    if dash: s.graphicalProperties.line.dashStyle=dash
    s.smooth=False
lc.series[0].marker.symbol="circle"; lc.series[0].marker.size=6
lc.series[0].marker.graphicalProperties.solidFill="C00000"; lc.series[0].marker.graphicalProperties.line.solidFill="C00000"
lc.set_categories(cats); lc.legend.position="b"; lc.display_blanks="gap"
lc.y_axis.number_format="0.00"; lc.y_axis.majorGridlines=None
lc.x_axis.delete=False; lc.y_axis.delete=False
lc.visible_cells_only=False
t.add_chart(lc,"A4")
bc=BarChart(); bc.title="Ümumi Cu bərpası"; bc.y_axis.title="Bərpa"; bc.height=8; bc.width=24
bc.add_data(Reference(g,min_col=15,min_row=4,max_row=N),titles_from_data=True)
bc.series[0].graphicalProperties.solidFill="2F5597"; bc.set_categories(cats); bc.legend=None
bc.y_axis.number_format="0%"; bc.y_axis.scaling.min=0; bc.y_axis.scaling.max=1
bc.x_axis.delete=False; bc.y_axis.delete=False; bc.y_axis.majorGridlines=None
bc.visible_cells_only=False
t.add_chart(bc,"A23")

# ================= DOZA =================
d=wb.create_sheet("Doza"); d.sheet_view.showGridLines=False
for col,w in zip("ABCDEFGHIJKL",[12,14,9,9,13,11,11,13,11,12,13,13]): d.column_dimensions[col].width=w
d.merge_cells("A1:L1"); c=d["A1"]; c.value="  Reagent doza kalkulyatoru"
c.font=font(size=16,bold=True,color="FFFFFF"); c.fill=fill(NAVY); c.alignment=L; d.row_dimensions[1].height=30
d.merge_cells("A2:L2"); c=d["A2"]
c.value="Doza (ml/t aktiv) = X ml / Y san × konsentrasiya/100 × 3600 / tonaj.   g/t = ml/t × sıxlıq.   Nümunə: 6 noyabr 00:17 sərfləri."
c.font=font(size=9,italic=True,color="595959")
d["A4"]="Tonaj"; d["A4"].font=font(bold=True); d["A4"].alignment=L
d["B4"]=450; inp(d["B4"],"0"); d["C4"]="t/saat"; d["C4"].font=font(color="595959")
d.merge_cells("E4:L4"); d["E4"]="Tonaj dəyişəndə bütün dozalar avtomatik yenilənir. 'Lazımi Y' sütunu: hədəf dozaya çatmaq üçün X ml neçə saniyəyə dolmalıdır."
d["E4"].font=font(size=9,italic=True,color="7F6000"); d["E4"].fill=fill(INP); d["E4"].alignment=L
d.row_dimensions[4].height=28
heads=[("Nöqtə","in"),("Reagent","in"),("X\nml","in"),("Y\nsan","in"),("Konsentrasiya\n%","in"),("Sıxlıq\ng/ml","in"),
       ("Sərf\nml/dəq","out"),("Doza\nml/t aktiv","out"),("Doza\ng/t","out"),("Hədəf doza\nml/t","in"),("Lazımi Y\nsan","out"),("Aktiv sərf\nL/gün","out")]
HR=6
for i,(h,tp) in enumerate(heads,1):
    c=d.cell(HR,i,h); c.font=font(bold=True,color="FFFFFF"); c.fill=fill(NAVY if tp=="in" else "375623"); c.alignment=C; c.border=BOX
d.row_dimensions[HR].height=32
rows=[("151 box","JAM",100,14,100,None),("151 box","AERO 5151",10,10,100,None),("151","SIPX",100,2,14,None),
      ("151","MIBC",50,4,100,0.81),("152","JAM",100,32,100,None),("152","SIPX",100,7,14,None),("153","MIBC",50,4,100,0.81)]
R0=HR+1; R1=HR+20
for r in range(R0,R1+1):
    for i in range(1,13):
        c=d.cell(r,i); c.border=BOX; c.alignment=C; c.font=font()
        c.fill = fill(INP) if heads[i-1][1]=="in" else fill(OUT)
    ok=f'OR(C{r}="",D{r}="",E{r}="")'
    d.cell(r,7,f'=IF({ok},"",C{r}/D{r}*60)').number_format="0.0"
    d.cell(r,8,f'=IF({ok},"",C{r}/D{r}*E{r}/100*3600/$B$4)').number_format="0.0"
    d.cell(r,9,f'=IF(OR(H{r}="",F{r}=""),"",H{r}*F{r})').number_format="0.0"
    d.cell(r,11,f'=IF(OR(J{r}="",C{r}="",E{r}=""),"",C{r}*E{r}/100*3600/($B$4*J{r}))').number_format="0.0"
    d.cell(r,12,f'=IF({ok},"",C{r}/D{r}*E{r}/100*86400/1000)').number_format="0.0"
for j,row in enumerate(rows):
    for i,v in enumerate(row,1):
        if v is not None:
            c=d.cell(R0+j,i,v); c.font=font(color="0000FF",bold=True)
# summary
S=R1+2
section(d,S,"Reagent üzrə cəmi doza",1,4)
for i,h in enumerate(["Reagent","ml/t aktiv","g/t"],1):
    c=d.cell(S+1,i,h); c.font=font(bold=True,color="FFFFFF"); c.fill=fill("44546A"); c.alignment=C; c.border=BOX
for j,rg in enumerate(["JAM","SIPX","AERO 5151","MIBC","PAX","SMBS"]):
    r=S+2+j
    d.cell(r,1,rg).font=font(bold=True); d.cell(r,1).border=BOX; d.cell(r,1).alignment=C
    c=d.cell(r,2,f'=SUMIF($B${R0}:$B${R1},A{r},$H${R0}:$H${R1})'); out(c,"0.0")
    c=d.cell(r,3,f'=SUMIF($B${R0}:$B${R1},A{r},$I${R0}:$I${R1})'); out(c,"0.0")
r=S+2+6
d.cell(r,1,"Kollektor cəmi").font=font(bold=True); d.cell(r,1).border=BOX
c=d.cell(r,2,f"=B{S+2}+B{S+3}+B{S+6}"); out(c,"0.0",bold=True)
d.cell(r,3,"JAM+SIPX+PAX").font=font(size=9,italic=True,color="595959")
dvr=DataValidation(type="list",formula1='"JAM,SIPX,AERO 5151,MIBC,PAX,SMBS,Lime,MX-3754,F-521"',allow_blank=True)
d.add_data_validation(dvr); dvr.add(f"B{R0}:B{R1}")
dvn=DataValidation(type="decimal",operator="greaterThan",formula1="0",showErrorMessage=True,error="0-dan böyük rəqəm")
d.add_data_validation(dvn); dvn.add(f"C{R0}:F{R1}"); dvn.add(f"J{R0}:J{R1}"); dvn.add("B4")
d.freeze_panes=f"A{R0}"

# ================= TƏLİMAT =================
h=wb.create_sheet("Təlimat"); h.sheet_view.showGridLines=False
h.column_dimensions["A"].width=4; h.column_dimensions["B"].width=28; h.column_dimensions["C"].width=95
h.merge_cells("A1:C1"); c=h["A1"]; c.value="  Təlimat və iş qaydaları"
c.font=font(size=16,bold=True,color="FFFFFF"); c.fill=fill(NAVY); c.alignment=L; h.row_dimensions[1].height=30
r=3
def block(name,items):
    global r
    section(h,r,name,1,3); r+=1
    for a,b in items:
        h.cell(r,2,a).font=font(bold=True); h.cell(r,2).alignment=Alignment(vertical="top",wrap_text=True)
        h.cell(r,3,b).font=font(); h.cell(r,3).alignment=Alignment(vertical="top",wrap_text=True)
        h.row_dimensions[r].height=max(16,15*(len(b)//105+1))
        r+=1
    r+=1
block("Rəng açarı",[("Sarı xana, göy rəqəm","Daxil edilən dəyər — yalnız bunları dəyiş"),
                    ("Yaşıl xana, qara rəqəm","Düsturla hesablanır — toxunma"),
                    ("Qırmızı rəqəm (Gündəlik)","Diqqət: S/Cu > 1.3 (pirit), conc Si > 3, cl. scav tail > 0.4, pH > 11.8"),
                    ("Vəziyyət rəngləri","ƏLA (yaşıl) · NORMAL (göy) · DİQQƏT (sarı) · PİS (qırmızı)")])
block("Vərəqlər",[("Kalkulyator","Bir analiz dəsti üçün: oksid/sulfid bölgüsü, mass pull, bərpa, hədəf tail zolağı, vəziyyət, Cu itkisi t/gün, istənilən tail üçün reallıq yoxlaması. Fərziyyələr (bölmə 2) bütün faylda istifadə olunur."),
                  ("Gündəlik","Hər analiz dəsti bir sətir. A–M və Qeyd sütunlarını doldur; qalanı avtomatikdir. CuOx yoxdursa hədəf fərziyyə ilə təxmin edilir (CuOx mənbəyi = Təxmin)."),
                  ("Trend","Total tail Cu-nun hədəf zolağı (aşağı/yuxarı) ilə müqayisəsi və bərpa qrafiki — Gündəlikdən avtomatik."),
                  ("Doza","X ml / Y san ölçülərindən ml/t və g/t. Tonajı B4-də dəyiş. Hədəf doza yazsan, lazımi Y saniyəni verir.")])
block("Düsturlar",[("Mass pull","(f − t) / (c − t)"),("Ümumi bərpa","c(f − t) / f(c − t)"),
                   ("Sulfid bərpası","1 − (tail Cu − tail CuOx) × (1 − mass pull) / (feed Cu − feed CuOx)"),
                   ("Hədəf tail","Tail CuOx + feed sulfid Cu × (1 − sulfid bərpası) / (1 − mass pull) + cleaner scav əlavəsi"),
                   ("Doza","X/Y × konsentrasiya/100 × 3600 / tonaj  (500 t/saat: ×7.2; 450 t/saat: ×8). MIBC: ml/t × 0.81 = g/t")])
block("İş qaydaları",[("1","Bir dəfəyə bir dəyişiklik, sonra 1–2 analiz (2–4 saat) gözlə."),
                      ("2","Hər dəyişiklikdən sonra tail Cu, conc Cu, S/Cu, Fe/Cu, Si izlə; mümkünsə Mg/MgO və CuOx."),
                      ("3","Lime və SMBS-i eyni anda dəyişmə (SMBS pH-ı salır)."),
                      ("4","Tonaj dəyişəndə sərfləri yenidən hesabla (Doza vərəqi)."),
                      ("5","Silikat depressantını yalnız conc Si yüksələndə düşün; talk üçün CMC/quar — əvvəlcə Mg + XRD ilə təsdiqlə."),
                      ("6","Hər növbədə eyni vaxtda tam dəst: feed, rougher tail, total tail, cleaner scavenger tail, conc.")])
block("Fərziyyələr və mənbə",[("Lab dataları","5–9 noyabr 2026 analizləri (istifadəçi tərəfindən verilib)."),
                              ("Sulfid bərpası 85–90%","Ümumi təcrübə dəyəri — zavod tarixçəsi ilə dəqiqləşdir (Kalkulyator B17–B18)."),
                              ("Cleaner scav 0.02–0.03","Təxmini — cleaner scavenger tail analizləri ilə dəqiqləşdir."),
                              ("Oksid payı 15.77%, oksid bərpası 27%","8 noyabr 23:00 oksid analizindən.")])

wb.move_sheet("Təlimat",offset=-4)
wb.active=1
for ws in wb.worksheets: ws.sheet_properties.tabColor={"Təlimat":"7F7F7F","Kalkulyator":"1F3864","Gündəlik":"375623","Trend":"C00000","Doza":"BF8F00"}[ws.title]
import os
out_path=os.path.join(os.path.dirname(os.path.abspath(__file__)),"Tail_hedef_kalkulyatoru.xlsx")
wb.save(out_path); print(out_path)
