import yfinance as yf
import pathlib
import os
import xlsxwriter
import datetime
import decimal

class DCF():
    def __init__(self,ticker,rf=1.48,name=''):
        self.ticker=yf.Ticker(ticker)
        self.rf=rf
        self.path=os.path.abspath(os.getcwd())
        self.desktop = os.path.join(os.path.join(os.path.expanduser('~')), 'Desktop') 
        self.today=datetime.date.today().year
        if name=='':
            self.name=self.desktop+"\\DCF_"+str(self)+'.xlsx'
        else:
            self.name=self.desktop+"\\"+self.name+'.xlsx'
        self.default()
        
    def __repr__(self):
      
        return  self.ticker.info['symbol']+'-US'
    def drange(self,x, y, jump):
          while x < y:
            yield float(x)
            x += decimal.Decimal(jump)
    def default(self):
        self.workbook=xlsxwriter.Workbook(self.name)
        self.sheet1 = self.workbook.add_worksheet('DCF')        # Defaults to Sheet1.
        self.sheet2 = self.workbook.add_worksheet('Prediction')  # Data.
        self.sheet3 = self.workbook.add_worksheet('IRR')
        try:
            pYear=self.choose_year()
        except :
            
            pYear=self.today-9
        
        
        
        ### format
        arial=self.workbook.add_format({'font_name':'Arial','font_size':11})
        blue_bold=self.workbook.add_format({'bold': True, 'bg_color': '#5B9BD5'})
        bold=self.workbook.add_format({'bold': True})
        blue=self.workbook.add_format({'bg_color': '#5B9BD5'})
        yellow=self.workbook.add_format({'bg_color': '#FFFF00'})
        grey=self.workbook.add_format({'bg_color': '#C8C8C8'})
        grey_bold=self.workbook.add_format({'bold': True, 'bg_color': '#C8C8C8'})
        orange_bold=self.workbook.add_format({'bold': True, 'bg_color': '#FFC000'})
        orange= self.workbook.add_format({'bg_color': '#FFC000'})                              
        green=  self.workbook.add_format({'bg_color': '#92D050'}) 
        white_bold=self.workbook.add_format({'bold': True, 'bg_color': '#FFFFFF','font_name':'Arial','font_size':11})
        white=  self.workbook.add_format({'bg_color': '#FFFFFF','font_name':'Arial','font_size':11})                             
        currency = self.workbook.add_format({'num_format': 8})                       
        percentage=  self.workbook.add_format({'num_format': 10})                             
        blue_currency= self.workbook.add_format({'bg_color': '#5B9BD5','num_format': 8}) 
        yellow_percentage=self.workbook.add_format({'bg_color': '#FFFF00','num_format': 10})
        num=self.workbook.add_format({'num_format': 1}) 
        num_bold=self.workbook.add_format({'num_format': 1,'bold': True}) 
        # sheet 1
        ## default data
        
        self.sheet1.set_row(4,15,bold)
        self.sheet1.set_column('C:{}'.format(self.add('C',self.today-pYear+7)), 25,arial)
#         self.sheet1.set_column('C3:{}3'.format(self.add('C',self.today-pYear+7)),15,grey_bold )
        self.sheet1.set_column('B:B', 50,arial)
        self.sheet2.set_column('A:AA', 10,white)
        self.sheet2.set_column('D:I', 25,white)
        self.sheet3.set_column('B:J', 25,arial)

        self.sheet1.write("B1",'Ticker',blue_bold)
        self.sheet1.write("C1",str(self),blue_bold)
        self.sheet1.write("D1",'=FDS(C1,\"FG_PRICE(NOW)\")',blue_bold)
        
        self.sheet1.write("B5","Sales")
        self.sheet1.write("B6",'Revenue Growth')
        self.sheet1.write("B9",'Net Income ')
        self.sheet1.write("B10",'Profit Margin')
        
        self.sheet1.write("B13","Assumed FCFE / Net Income")
        self.sheet1.write("B14",'Free Cash Flow')
        self.sheet1.write("B15",'FactSet estimate ')
        self.sheet1.write("B16",'Estimate selling price FV')
        self.sheet1.write("B17",'Estimate selling price PV')


        self.sheet1.write("B20","Beta")
        self.sheet1.write("B21",'Rf')
        self.sheet1.write("B22",'E(Rm)')
        self.sheet1.write("B24",'CAPM Required Return')
        
        self.sheet1.write("B26","Shares Outstanding")
        self.sheet1.write("B29",'Intrinsic Value of Equity (User estimate)')
        self.sheet1.write("B30",'Intrinsic Value per Share')
        self.sheet1.write("B32",'Intrinsic Value of Equity (FactSet estimate)',blue)
        self.sheet1.write("B33",'Intrinsic Value per Share',blue)
        self.sheet1.write("B36",'Sensitivity Analysis',bold,)
        self.sheet1.write("B38",'Required return')
        
        
        
        self.sheet1.write_formula("C20",'=FDS(C1,"XP_BETA_PR(NOW,NOW,,""6M"")")',yellow)
        self.sheet1.write_formula("C21",'={}%'.format(self.rf),yellow_percentage)
        self.sheet1.write("C22","=8%",yellow_percentage)
        self.sheet1.write_formula("C24","=C21+C20*(C22-C21)",yellow_percentage)
        self.sheet1.write_formula("C26",'=FDS(C1,"FF_COM_SHS_OUT(ANN_R,NOW,NOW)")')
        self.sheet1.write_formula("C30","=C29/C26")
        self.sheet1.write_formula("C33","=C32/C26",blue_bold)
        self.sheet1.write("C38","Intrinsic Value")
        self.sheet1.write_formula("C17","=C16/(1+C24)^5")
        self.workbook.define_name('beta', "=DCF!$C$20")

        self.sheet1.write("D20","*Obtained from FactSet.")
        self.sheet1.write("D21","*Obtained from Treasury.gov. Used 10-Yr Treasury Yield")        
        self.sheet1.write("D22","*User preferences, default 8%")
        self.sheet1.write("D26","*Obtained from Annual Report")
        
        
        self.sheet1.write("E38","Perpetual Growth Rate")
        self.sheet1.write("F38","Intrinsic Value")
                                       
                                       
#         self.sheet1.write("H38","Beta")
#         self.sheet1.write("I38","Intrinsic Value")
        
        
        self.sheet1.write("O21","Assumptions")
        self.sheet1.write("O22","*1. Revenue estimates for {}-{} taken from FactSet's report of analyst estimates. If Factset doesn't have the estimate take from user input".format(self.today,self.today+5))
        self.sheet1.write("O23","*2. Stable revenue growth estimate for {} and beyond is assumed 1.5%".format(self.today+6))
        self.sheet1.write("O24","*3. Assumed FCFE will be 75% of net income for default")

        
        ## function data


        
        col='C'
        row=4
        
        for _ in range(pYear,self.today+1):
            self.sheet1.write_blank(col+str(row-1),None,grey)
            self.sheet1.write(col+str(row),_)
            row+=1
            self.sheet1.write_formula(col+str(row),'=IF(ISNA(FDS($C$1,"FF_SALES(ANN_R,"&{}&"-1,"&{}&"-1,,,USD)")),0,FDS($C$1,"FF_SALES(ANN_R,"&{}&"-1,"&{}&"-1,,,USD)"))'.format(col+str(row-1),col+str(row-1),col+str(row-1),col+str(row-1)),num_bold)
            row+=1
            self.sheet1.write_formula(col+str(row),'=IFERROR({}/{}-1,0)'.format(col+str(row-1),self.subtract(col)+str(row-1)),percentage)
            row+=3
            self.sheet1.write_formula(col+str(row),'=IF(ISNA(FDS($C$1,"FF_NET_INC(ANN_R,"&{}&"-1,"&{}&"-1,,,USD)")),0,FDS($C$1,"FF_NET_INC(ANN_R,"&{}&"-1,"&{}&"-1,,,USD)"))'.format(col+str(row-5),col+str(row-5),col+str(row-5),col+str(row-5)),num)
            row+=1
            self.sheet1.write_formula(col+str(row),'=IFERROR({}/{},0)'.format(col+str(row-1),col+str(row-5)),percentage)                          
            col=self.add(col)
            row-=6
        
        self.sheet1.write(col+str(row+4),"(Assumed 25% Profit Margin on Forecasted Sales)")
        self.sheet1.write(col+str(row+8),"(Assumed FCFE is 75% of Net Income)")
        
        self.sheet1.write_formula("C29","=NPV(C24,{}14,{}14,{}14,{}14,{}14+{}14)".format(col,self.add(col),self.add(col,2),self.add(col,3),self.add(col,4),self.add(col,5)),currency)
        
        self.sheet1.write_formula("C32","=NPV(C24,{}15,{}15,{}15,{}15,{}15+{}15)".format(col,self.add(col),self.add(col,2),self.add(col,3),self.add(col,4),self.add(col,5)),blue_currency)
        #M4
        
        c1="B"
        r1=40

        for _ in range(6,14):
            self.sheet1.write(c1+str(r1),"={}%".format(_),percentage)
            self.sheet1.write_formula(self.add(c1)+str(r1),"=NPV({},{}15,{}15,{}15,{}15,{}15+{}15)/C26".format(c1+str(r1),col,self.add(col),self.add(col,2),self.add(col,3),self.add(col,4),self.add(col,5)))
            
            
            r1+=1
        c1="E"
        r1=40
        for _ in list(self.drange(-3,3,'0.5')):
            self.sheet1.write(c1+str(r1),"={}%".format(_),percentage)
            self.sheet1.write_formula(self.add(c1)+str(r1),"=NPV(C24,{}15,{}15,{}15,{}15,{}15+{}15*(1+{})/(C24-{}))/C26".format(col,self.add(col),self.add(col,2),self.add(col,3),self.add(col,4),self.add(col,4),c1+str(r1),c1+str(r1)))

            r1+=1
            
#         c1="H"
#         r1=40
#         for _ in range(5):
#             self.sheet1.write_formula(c1+str(r1),"=beta-0.4+0.1*{}".format("_"))
#             self.sheet1.write_formula(self.add(c1)+str(r1),"=NPV(C24,{}15,{}15,{}15,{}15+{}15+{}15*(1+{})/(C24-{}))".format(col,self.add(col),self.add(col,2),self.add(col,3),self.add(col,4),self.add(col,4),c1+str(r1),c1+str(r1)))

#             r1+=1
#         for _ in range(1,5):
#             self.sheet1.write_formula(c1+str(r1),"=beta-0.1*{}".format("_"))
#             r1+=1
        
        
        
        
        
        
        
        
        self.workbook.define_name('ni', "=DCF!${}${}".format(self.subtract(col),row+5))
        
        
        cc=col
        
        
        #M4
        for _ in range(self.today+1,self.today+6):
            self.sheet1.write(col+str(row),_)
            
            row-=1
            self.sheet1.write(col+str(row),"***Est",grey_bold)
            row+=2
            self.sheet1.write_formula(col+str(row),'=IF(ISNA(FDS($C$1,"FE_ESTIMATE(SALES,MEAN,ANN_ROLL,"&{}4&"-1,NOW,NOW,,'')")),{}*(1+pRate),FDS($C$1,"FE_ESTIMATE(SALES,MEAN,ANN_ROLL,"&{}4&"-1,NOW,NOW,,'')"))'.format(col,self.subtract(col)+str(row),col),num_bold)
            row+=1
            self.sheet1.write_formula(col+str(row),'={}/{}-1'.format(col+str(row-1),self.subtract(col)+str(row-1)),percentage)
            row+=3
            self.sheet1.write_formula(col+str(row),'={}*{}'.format(col+str(row-4),col+str(row+1)))
            row+=1
            self.sheet1.write(col+str(row),'=25%',yellow_percentage)
            row+=3
            self.sheet1.write(col+str(row),'=75%',yellow_percentage)
            row+=1
            self.sheet1.write_formula(col+str(row),'={}*{}'.format(col+str(row-5),col+str(row-1)))
            row+=1
            self.sheet1.write_formula(col+str(row),'=IF(ISNA(FDS(DCF!$C$1,"FE_ESTIMATE(FCF,MEAN,ANN_ROLL,"&{}4&",NOW,NOW,,'')")),{}14,FDS(DCF!$C$1,"FE_ESTIMATE(FCF,MEAN,ANN_ROLL,"&{}4&",NOW,NOW,,'')"))'.format(col,col,col))
            row-=11
            
            col=self.add(col)
         
        #R4
        self.sheet1.write_formula("C16","={}14/fshares".format(col))
        
        self.sheet1.write(col+str(row),"{} and beyond".format(self.today+6))
        self.sheet1.write(col+str(row+9),"{} Terminal value".format(self.today+6))
        self.sheet1.write_formula(col+str(row+10),"={}14*(1+{}6)/(C24-{}6)".format(self.subtract(col),col,col))
        self.sheet1.write_formula(col+str(row+11),"={}15*(1+{}6)/(C24-{}6)".format(self.subtract(col),col,col))
        
        
        row-=1
        self.sheet1.write(col+str(row),"***Est",grey_bold)

        self.sheet1.write(col+str(row-1),"Assumed eternal growth = 1.5%")
        row+=3
        self.sheet1.write(col+str(row),"=1.5%",percentage)
        row-=1
        col=self.add(col)
        self.sheet1.write(col+str(row),"*** For user prediction",orange_bold)
        row+=5
        self.sheet1.write(col+str(row),"*** For user prediction",orange_bold)
        row+=3
        self.sheet1.write(col+str(row),"*** For user prediction",orange_bold)
        
        

        
        # sheet 2
        
        col='D'
        row=5
        self.sheet2.write(col+str(row),"Year")
        col=self.add(col)
        self.sheet2.write(col+str(row),"Equity(M)")
        col=self.add(col)
        self.sheet2.write(col+str(row),"Share(M)")
        col=self.add(col)
        self.sheet2.write(col+str(row),"Equity/Share")
        col=self.add(col)
        self.sheet2.write(col+str(row),"BookValue/share")
        col=self.add(col)
        self.sheet2.write(col+str(row),"Modify")
        
        col=self.subtract(col,5)
        row+=1
        
        for _ in range(pYear-1,self.today):
            self.sheet2.write(col+str(row),_)
            
            
            self.sheet2.write_formula(self.add(col)+str(row),'=IF(ISNA(FDS(DCF!$C$1,"FF_COM_EQ(ANN_R,"&{}&","&{}&",,,USD)")),0,FDS(DCF!$C$1,"FF_COM_EQ(ANN_R,"&{}&","&{}&",,,USD)"))'.format(col+str(row),col+str(row),col+str(row),col+str(row)))
            self.sheet2.write_formula(self.add(col,2)+str(row),'=IF(ISNA(FDS(DCF!$C$1,"FF_COM_SHS_OUT(ANN_R,"&{}&","&{}&")")),0,FDS(DCF!$C$1,"FF_COM_SHS_OUT(ANN_R,"&{}&","&{}&")"))'.format(col+str(row),col+str(row),col+str(row),col+str(row)))
            self.sheet2.write_formula(self.add(col,3)+str(row),'={}/{}'.format(self.add(col)+str(row),self.add(col,2)+str(row)))
            self.sheet2.write_formula(self.add(col,4)+str(row),'=FDS(DCF!$C$1,"FF_BPS(ANN_R,"&{}&","&{}&",,,USD)")'.format(col+str(row),col+str(row)))
            self.sheet2.write_formula(self.add(col,5)+str(row),'=IF(ISNA(IF({}<0,EXP({}),{})),0.1,IF({}<0,EXP({}),{}))'.format(self.add(col,4)+str(row),self.add(col,4)+str(row),self.add(col,4)+str(row),self.add(col,4)+str(row),self.add(col,4)+str(row),self.add(col,4)+str(row)))
            row+=1
        
        self.sheet2.write("D{}".format(row+7),"Growth")
        self.sheet2.write_formula("E{}".format(row+7),'=(I{}/I6)^(1/{})-1'.format(row-1,row-6))
        
        self.workbook.define_name('pRate', "=Prediction!$E${}".format(row+7))
        
        
        
        #sheet 3
        
        col='B'
        row=3
        self.sheet3.write(col+str(row),"Time",bold)
        col=self.add(col)
        self.sheet3.write(col+str(row),"Year",bold)
        col=self.add(col)
        self.sheet3.write(col+str(row),"Computed FCFE",bold)
        col=self.add(col)
        self.sheet3.write(col+str(row),"Number Shares",bold)
        col=self.add(col)
        self.sheet3.write(col+str(row),"FCFE/Share",bold)
        col=self.add(col)
        self.sheet3.write(col+str(row),"Investor CF",bold)
        
        col=self.subtract(col,5)
        row+=2
        r=self.today-pYear
        for _ in range(r,0,-1):
            self.sheet3.write(col+str(row),"=-{}".format(_))
            self.sheet3.write(self.add(col)+str(row),pYear+r-_)
            self.sheet3.write_formula(self.add(col,2)+str(row),'=FDS(DCF!$C$1,"FF_FREE_CF(ANN_R_FCF,"&{}&"-1,"&{}&"-1,,,USD)")'.format(self.add(col)+str(row),self.add(col)+str(row)))
            self.sheet3.write_formula(self.add(col,3)+str(row),'=IF(ISNA(FDS(DCF!$C$1,"FF_COM_SHS_OUT(ANN_R,"&{}&"-1,"&{}&"-1)")),{},FDS(DCF!$C$1,"FF_COM_SHS_OUT(ANN_R,"&{}&"-1,"&{}&"-1)"))'.format(self.add(col)+str(row),self.add(col)+str(row),self.add(col,3)+str(row+1),self.add(col)+str(row),self.add(col)+str(row)))
            self.sheet3.write_formula(self.add(col,4)+str(row),'={}/{}'.format(self.add(col,2)+str(row),self.add(col,3)+str(row)))
            row+=1
                                  
        self.sheet3.write(col+str(row),0)                          
        self.sheet3.write(self.add(col)+str(row),self.today)
        self.sheet3.write_formula(self.add(col,2)+str(row),'=IF(ISNA(FDS(DCF!$C$1,"FF_FREE_CF(ANN_R_FCF,"&{}&"-1,"&{}&"-1,,,USD)")),FDS(DCF!$C$1,"FE_ESTIMATE(FCF,MEAN,ANN_ROLL,"&DCF!{}4&",NOW,NOW,,'')"),FDS(DCF!$C$1,"FF_FREE_CF(ANN_R_FCF,"&{}&"-1,"&{}&"-1,,,USD)"))'.format(self.add(col)+str(row),self.add(col)+str(row),self.subtract(cc),self.add(col)+str(row),self.add(col)+str(row)))
        self.sheet3.write_formula(self.add(col,3)+str(row),'=FDS(DCF!$C$1,"FF_COM_SHS_OUT(ANN_R,"&{}&"-1,"&{}&"-1)")'.format(self.add(col)+str(row),self.add(col)+str(row)))
        self.sheet3.write_formula(self.add(col,4)+str(row),'={}/{}'.format(self.add(col,2)+str(row),self.add(col,3)+str(row)))
        self.sheet3.write_formula(self.add(col,5)+str(row),'=-FDS(DCF!C1,"FG_PRICE(NOW,NOW)")')
        
        
        self.sheet3.write_formula('J6','=(E{}/E5)^(1/{})-1'.format(row,r))
        self.workbook.define_name('fshares', "=IRR!${}${}".format(self.add(col,3),row+5))
        row+=1
        for _ in range(1,6):
            self.sheet3.write(col+str(row),_)                          
            self.sheet3.write(self.add(col)+str(row),self.today+_)
            self.sheet3.write_formula(self.add(col,2)+str(row),'=DCF!{}15'.format(cc))
            cc=self.add(cc)
            self.sheet3.write_formula(self.add(col,3)+str(row),'={}*(1+J$6)^{}'.format(self.add(col,3)+str(row-1),1))
            self.sheet3.write_formula(self.add(col,4)+str(row),'={}/{}'.format(self.add(col,2)+str(row),self.add(col,3)+str(row)))
            self.sheet3.write_formula(self.add(col,5)+str(row),'={}'.format(self.add(col,4)+str(row)))
            row+=1
        
        self.sheet3.write_formula(self.add(col,2)+str(row-1),'=DCF!{}15+DCF!{}15'.format(self.subtract(cc),cc))
        row+=3                      
        col=self.add(col,3)                          
        self.sheet3.write(col+str(row),"Internal Rate of Return")  
        self.sheet3.write_formula(self.add(col,2)+str(row),"=IRR(G{}:G{})".format(row-9,row-4),percentage)  
        self.sheet3.write(self.add(col,3)+str(row),"<-- Forecasted Return from buying at current market price given projected cash flows")  
                  
                                  
                                  
                                  
                                  
        
        ### define
        
        
        
        
        
        

        self.save()
        
    
    def choose_year(self):
        d=datetime.date(year=self.today-15,month=1,day=1)
        a=self.ticker.history(period='15y',interval='3mo',start=d)
        k=a.to_dict()['Close']
        for _ in k:
          i=_.to_pydatetime()

          break
        if i.year>d.year:
            return i.year-2
        else:
            return d.year
    def check_ticker(self,ticker):
        
        
        return len(yf.Ticker(ticker).info)>2
    def add(self,a,n=1):
        b=a
        for _ in range(n):
            b=chr(ord(b)+1)
        return b
    def subtract(self,a,n=1):
        b=a
        for _ in range(n):
            b=chr(ord(b)-1)
        return b

    
    def save(self):
        self.workbook.close()