def dfs(coin,required,count=[0],s=[0]):
    if s[0]==required:
        count[0]+=1
        return count
    print(coin,s)
    for i in range(len(coin)):
        
        if coin[i]>0:
            coin[i]-=1
            dfs(coin,required,count,s)
            if i==0:
                
                dfs(coin,required,count,[s[0]+10])
                break
            elif i==1:
                dfs(coin,required,count,[s[0]+5])
                break
            elif i==2:
                dfs(coin,required,count,[s[0]+2])
                break
            elif i==3:
                dfs(coin,required,count,[s[0]+1])
                break
    return count






dfs([1,1,3,2],15,count=c)

















