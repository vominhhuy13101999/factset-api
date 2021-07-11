from dcf import DCF
import os
import yfinance as yf
import webbrowser


def check_ticker(ticker):
        
        
        return len(yf.Ticker(ticker).info)>5

if __name__ == "__main__":
    for _ in range(5):
        print("Please enter ticker\'s name: ")
        ticker=str(input())
        if check_ticker(ticker):
            break
        else:
            ticker=''
        
    if ticker=='':
        exit()
    else:
        print("Next we will ask for risk free rate. The default rate will be 1.48%. Would you want to consult Treasury yield for risk free rate data?[Y/n]")
        a=str(input())
        
        if a=='y'or a=='Y' or a=='\n' or a=='':
            webbrowser.open("https://www.treasury.gov/resource-center/data-chart-center/interest-rates/Pages/TextView.aspx?data=yield")

        
        print("please enter risk free rate(please put 1.48 for 1.48% or enter to pass): ")

        try:
            rf=float(input())
            
        except ValueError:
            rf=1.48
            
        
        if rf<-3 or rf>2:
            rf=1.48

        
        print("----------------loading-------------")
        rf=float(input())
        DCF(ticker,rf)
        print("Done!-please check your Desktop")