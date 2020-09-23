<div align="center">

## CipherII


</div>

### Description

A StreamCipher encryption similar to the first 'Cipher' but now it can encrypt any text or binary file.
 
### More Info
 
PlainText(the text to be encrypted/decrypted), Secret(the password)

There is only one function to Encrypt/Decrypt the strings. It is to this point stable that you can encrypt the same string multiple times using different passwords every time and then decrypt in the reverse order.

Encrypted/Decrypted string

I tested the code in VB6.0 but I mainly used code that's been part of VB for a while so it should work with ver. 3 and up. One side effect is the cipher speed. It's still a little slow but I'm working on an 8, 16, and 32 cipher code which would hopefully be fast. Warning!! Keep the strings your working with in memory rather than in an object such as a textbox because there would be loss of data when a string with ascii < 32 is used in a textbox.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[John Cui](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-cui.md)
**Level**          |Unknown
**User Rating**    |3.2 (16 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Encryption](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/encryption__1-48.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/john-cui-cipherii__1-1310/archive/master.zip)





### Source Code

```
Public Function Cipher(PlainText, Secret)
Dim a, b, c
Dim pTb, cTb, cT
For i = 1 To Len(PlainText)
  pseudoi = i Mod Len(Secret)
  If pseudoi = 0 Then pseudoi = 1
  a = Mid(Secret, pseudoi, 1)
  b = Mid(Secret, pseudoi + 1, 1)
  c = Asc(a) Xor Asc(b)
  pTb = Mid(PlainText, i, 1)
  cTb = c Xor Asc(pTb)
  cT = cT + Chr(cTb)
  Form1.Label1.Caption = i
  DoEvents
Next i
EnCipher = cT
End Function
```

