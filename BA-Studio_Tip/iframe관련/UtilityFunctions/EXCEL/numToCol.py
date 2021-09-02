def numToCol(n) :
    string = ""
    n = n+1
    
    while n > 0 :
       
        n, remainder = divmod(n, 26)
        if(remainder is 0) :
            n = n-1
            string = 'Z' + string
        
        else :
            string = chr(ord('A') + remainder -1) + string

    return string

#Driver code
print(numToCol(727))
