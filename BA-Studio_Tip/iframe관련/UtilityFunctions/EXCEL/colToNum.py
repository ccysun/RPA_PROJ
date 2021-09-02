def colToNum(colStr):
    """ Convert base26 column string to number. """
    expn = 0
    colNum = 0
    for char in reversed(colStr):
        colNum += (ord(char) - ord('A') + 1) * (26 ** expn)
        expn += 1

    return colNum-1

#Driver code
print(colToNum('T'))
