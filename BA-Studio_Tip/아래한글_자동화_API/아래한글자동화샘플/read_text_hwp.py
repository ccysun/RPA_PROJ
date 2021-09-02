import olefile

f = olefile.OleFileIO('C:\\temp2\\아래한글예제\\사령장.hwp')
encoded_text = f.openstream('PrvText').read()
decoded_text = encoded_text.decode('UTF-16')
print(decoded_text)
