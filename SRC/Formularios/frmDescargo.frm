VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmDescargo 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descargos Almacén"
   ClientHeight    =   7155
   ClientLeft      =   5700
   ClientTop       =   6000
   ClientWidth     =   11070
   Icon            =   "frmDescargo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   11070
   Begin VB.Frame Frame2 
      Height          =   5805
      Left            =   4440
      TabIndex        =   10
      Top             =   480
      Width           =   6585
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   5475
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   6300
         _ExtentX        =   11113
         _ExtentY        =   9657
         _LayoutType     =   4
         _RowHeight      =   19
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   16
         Columns(0)._MaxComboItems=   5
         Columns(0).ValueItems(0)._DefaultItem=   0
         Columns(0).ValueItems(0).Value=   "0"
         Columns(0).ValueItems(0).Value.vt=   8
         Columns(0).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
         Columns(0).ValueItems(0).DisplayValue(0)=   "bHQAADYDAABCTTYDAAAAAAAANgAAACgAAAAQAAAAEAAAAAEAGAAAAAAAAAMAAAAAAAAAAAAAAAAA"
         Columns(0).ValueItems(0).DisplayValue(1)=   "AAAAAAD///////////8An+IAneAAmt4Aldv///////8Ag8wAfsgAecQAdcD/////////////////"
         Columns(0).ValueItems(0).DisplayValue(2)=   "//8ApecAouVRxvJEvu8Amt7///////8AitFgxvI8quIAesQAdcH///////////////8AqOkzu++C"
         Columns(0).ValueItems(0).DisplayValue(3)=   "4P532vwAneEAm98Al9wkpuJq1/9n0PkZktQAe8X///////////////////8AqOo+wfGF4v86u+4A"
         Columns(0).ValueItems(0).DisplayValue(4)=   "nuEAm99NwvFs2P82ruUAhs7///////////////////////////8AqepIxfNy2PoAoeRSxvNz2/9B"
         Columns(0).ValueItems(0).DisplayValue(5)=   "u+0Ak9n///////////////////////////////8ArOwAqep93v1j1fxn2v9e0PkAnOAAmd7/////"
         Columns(0).ValueItems(0).DisplayValue(6)=   "//////////8At/YAtfQAs/IAsfAAru8ArO1d0vph2f9b1/9c0PoAn+IAneAAmt4AldsAj9YAidEA"
         Columns(0).ValueItems(0).DisplayValue(7)=   "uvh63/2k7P+B4P112/ty2vt54P9h2v9c2P9r2/9v1fpl0Ph11/uF3v9fy/UAkNYAvPkixPkAuPYA"
         Columns(0).ValueItems(0).DisplayValue(8)=   "tvUAtPMAsvFn2fxk2/9g2f9f0/sApecAo+UAoOMAneEaqOUAl9wAvvsAvPn///////8AtvUAtPOD"
         Columns(0).ValueItems(0).DisplayValue(9)=   "4v5w3v5x3f9m1fsAqOoApuj///////8AnuEAm9////////////////8AufdN0PqJ5f4AsvJY0PmB"
         Columns(0).ValueItems(0).DisplayValue(10)=   "4f9Gx/UAqer///////////////////////////8AvfpE0PuR6P9X0/oAtfQAs/JW0Pl+4P86wvQA"
         Columns(0).ValueItems(0).DisplayValue(11)=   "qer///////////////////8AwP0tzPyH5f6R6f8VwPgAt/YAtfQpwfWA4f9y2fwArO0Aquv/////"
         Columns(0).ValueItems(0).DisplayValue(12)=   "//////////8Awv4iyf1n3P554P4Au/n///////8AtvR63f1RzvkAr+8ArO3/////////////////"
         Columns(0).ValueItems(0).DisplayValue(13)=   "//8Awv4Awf0Av/wAvvv///////8AuPYAtvUAtPMAsvH/////////////////////////////////"
         Columns(0).ValueItems(0).DisplayValue(14)=   "//////////////////////////////////////////8="
         Columns(0).ValueItems(0).DisplayValue.vt=   9
         Columns(0).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(0).ValueItems(1)._DefaultItem=   0
         Columns(0).ValueItems(1).Value=   "1"
         Columns(0).ValueItems(1).Value.vt=   8
         Columns(0).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
         Columns(0).ValueItems(1).DisplayValue(0)=   "bHQAADYDAABCTTYDAAAAAAAANgAAACgAAAAQAAAAEAAAAAEAGAAAAAAAAAMAAAAAAAAAAAAAAAAA"
         Columns(0).ValueItems(1).DisplayValue(1)=   "AAAAAAD////////9/v39/v39/v39/v39/v39/v39/v39/v39/v39/v39/v39/f3////////V1NZr"
         Columns(0).ValueItems(1).DisplayValue(2)=   "bWp/f3WCg3aCg3aCg3aCg3aCg3aCg3aCg3aCg3aCg3aCg3aBgXdYWlfT09X3+PWoqb4xO8coLsIq"
         Columns(0).ValueItems(1).DisplayValue(3)=   "MMIqMMIqMMIhKMEgJ8EqMMIqMMIqMMIpL8IrNMSipcZzdHHp6eNNT4ciL+wYINwaIt0aIt0SGtxr"
         Columns(0).ValueItems(1).DisplayValue(4)=   "ced+g+8QGdsaIt0aIt0ZId0VH+ZlarG8vbb///+fn5QxO78UHNwdJdodJdoTHNlzdtyEiOcSGtgd"
         Columns(0).ValueItems(1).DisplayValue(5)=   "JdodJdoXH9suOM6MjIf////////x8exeXns/SuwSGtgcJNkYINk3PttJT+EXH9kdJdobI9kcJ+R5"
         Columns(0).ValueItems(1).DisplayValue(6)=   "e6TPz8j///////////+6uqtMUa1ZX+sNFtcOFth0d9+Zne0QGNgcJNkVHN1ASsahoZP/////////"
         Columns(0).ValueItems(1).DisplayValue(7)=   "///////4+PVjZHuDie+Ch+0zOt6BhOOfovAFDtUSGtkkLt6CgpTn5uL////////////////////L"
         Columns(0).ValueItems(1).DisplayValue(8)=   "y787QZiZnfiFi+3IyfPT1flfZOVpb+1kacKnqJv///////////////////////////94eIBscuCJ"
         Columns(0).ValueItems(1).DisplayValue(9)=   "je/T0/Le3/pvc+pTW9uEhIv8/fr////////////////////////////b3NE/Q4uVmvrZ2vTn6Pty"
         Columns(0).ValueItems(1).DisplayValue(10)=   "ePFtcbixsaj///////////////////////////////////+Li4haYdPLzPPOz/dTW9ONjoj/////"
         Columns(0).ValueItems(1).DisplayValue(11)=   "///////////////////////////////////q6+RJS4Ccof+HjfZ7fanIx8L/////////////////"
         Columns(0).ValueItems(1).DisplayValue(12)=   "//////////////////////////+urqIrMZJvdseRkY7/////////////////////////////////"
         Columns(0).ValueItems(1).DisplayValue(13)=   "///////////////w8PDZ2da8vbbi4uL/////////////////////////////////////////////"
         Columns(0).ValueItems(1).DisplayValue(14)=   "//////////////////////////////////////////8="
         Columns(0).ValueItems(1).DisplayValue.vt=   9
         Columns(0).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(0).ValueItems(2)._DefaultItem=   0
         Columns(0).ValueItems(2).Value=   "2"
         Columns(0).ValueItems(2).Value.vt=   8
         Columns(0).ValueItems(2).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
         Columns(0).ValueItems(2).DisplayValue(0)=   "bHQAADYDAABCTTYDAAAAAAAANgAAACgAAAAQAAAAEAAAAAEAGAAAAAAAAAMAAAAAAAAAAAAAAAAA"
         Columns(0).ValueItems(2).DisplayValue(1)=   "AAAAAAD///////////////////+zvLNrjm4/cUQ2ZDtifWOloqX99v3/////////////////////"
         Columns(0).ValueItems(2).DisplayValue(2)=   "//////+yvrQfZCNLqFNszHNz34N52oJqxXBAlkcSRxaal5r///////////////////95m3sxjjg8"
         Columns(0).ValueItems(2).DisplayValue(3)=   "xEgMtyIXvjQpxkUhwz4OuicivzRhzWwZYx5saWv///////////+YsZk4mj8QsRwEqBELrRwArRAC"
         Columns(0).ValueItems(2).DisplayValue(4)=   "shoVtysPsSEJqhkAqAxHyFEUWxmcl5v//////v8ofi4atiYEpREGphMAnQD2+vb5+/gAnwAIqBUJ"
         Columns(0).ValueItems(2).DisplayValue(5)=   "qBYJpxUApAlQxloONhD++P2Mq40/uEoApAsFpREAnADh9OL///////+i3qYAnAAHphMIphQHphMU"
         Columns(0).ValueItems(2).DisplayValue(6)=   "syArgTKooadblV4ftiwAowwAogbg9OH////n8ej///////86uUQEpRAApAwCpA4ApQtCsUtmdGZH"
         Columns(0).ValueItems(2).DisplayValue(7)=   "kE0AqQsPqRqn4Kz///////9NtVWFxYr///////83t0Faw2M6uEMAoQVEwk87Vjw6i0BAwEun4K1+"
         Columns(0).ValueItems(2).DisplayValue(8)=   "xoSWxZllu21Fv04qrzXl6+X////p9+pbw2SP15aD04pIyVRDYkREjkmT3Zqa2qCW2p6E1It0znxm"
         Columns(0).ValueItems(2).DisplayValue(9)=   "x25Yw2FWvV7////////U79WC0oqi3qhmxG1rg21bk15/2Yeb2qGX2ZyV2JyK1JB/0YZ3zn9uzHWW"
         Columns(0).ValueItems(2).DisplayValue(10)=   "ypv///////+s4LCb3qIvmTe0urTK1so0nTyb3qKX2Z2X2ZyX2Z2S2JqQ15aP1pWD0IvM5c////+z"
         Columns(0).ValueItems(2).DisplayValue(11)=   "5Lh/3IcSWhb///////89fUFszXWg3aaX2Z2X2Z2X2Z2Y2Z2Y2Z6Z2p+Hz45msm2Z46AkiCqyvbP/"
         Columns(0).ValueItems(2).DisplayValue(12)=   "/////////v8ccSJszXWb3qKb2qCZ2Z6W2Z2X2Z6Z2Z6f3aWU4psulzZ3mXj/////////////////"
         Columns(0).ValueItems(2).DisplayValue(13)=   "/v88fUE0nTx/2YiM2pOg4aaX3p2F2Y1ixGsbeSGWsZf////////////////////////////K1cpb"
         Columns(0).ValueItems(2).DisplayValue(14)=   "k15DjkkwiDc4ij9PkFKGqYj//f////////////////8="
         Columns(0).ValueItems(2).DisplayValue.vt=   9
         Columns(0).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
         Columns(0).ValueItems(3)._DefaultItem=   -1
         Columns(0).ValueItems(3).Value=   "3"
         Columns(0).ValueItems(3).Value.vt=   8
         Columns(0).ValueItems(3).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
         Columns(0).ValueItems(3).DisplayValue(0)=   "bHQAADYDAABCTTYDAAAAAAAANgAAACgAAAAQAAAAEAAAAAEAGAAAAAAAAAMAAAAAAAAAAAAAAAAA"
         Columns(0).ValueItems(3).DisplayValue(1)=   "AAAAAAD////////////////////////////////////////////////////////////////+8u6r"
         Columns(0).ValueItems(3).DisplayValue(2)=   "nZibjIidjoqdjoqdjoqdjoqdjoqdjoqdjoqdjoqdjoqdjoqbjYivpJ/y8vJIrNJExfBOyPFNyPFM"
         Columns(0).ValueItems(3).DisplayValue(3)=   "x/FKx/FJyfRN1P9IyvVJx/FMyPFOyfFQyvFPz/gHd6G7sKtTw+c93P8n0v8z1/9A3P9T5P9i2/U0"
         Columns(0).ValueItems(3).DisplayValue(4)=   "YWlq1eta5P881v8my/8ZxP8jzP9Bu+PSxcCl2e1T2f0m0v8u1f832P9K5P82jqMAAAAxdIRU6f88"
         Columns(0).ValueItems(3).DisplayValue(5)=   "2P8v0P8cxf9G1v9Ulaz///3///9HuuJ25/9r4f9W3f882f865P9C0vRL6/8+2P841v8x1P842P80"
         Columns(0).ValueItems(3).DisplayValue(6)=   "q9Pi087////////J5/Ns0/PI9v+38P+w7/+Q+v8RExMy5v801v801f8r1v9X2P1zmKj/////////"
         Columns(0).ValueItems(3).DisplayValue(7)=   "//////9RvOK69f+68f+w8P+s+f+MdXBv5P8r1f8t1P9S4/8tk7j/8Or////////////////7/P06"
         Columns(0).ValueItems(3).DisplayValue(8)=   "vOfP+P+68f+x7v6Of3uk3ex85v8m1v9Bw+2wsrT///////////////////////96yeaf6f3J9v+2"
         Columns(0).ValueItems(3).DisplayValue(9)=   "4Oudko+k1eOp8f+D6v86i6n//vr///////////////////////////8rsuDX/v+/2uKpoqCszNO3"
         Columns(0).ValueItems(3).DisplayValue(10)=   "+v8jqNXez8v///////////////////////////////+03/By1/XW7fK0p6TB3+d34v9jk6f/////"
         Columns(0).ValueItems(3).DisplayValue(11)=   "//////////////////////////////////83st3U/f/U/v/N//8Qj7z86eP/////////////////"
         Columns(0).ValueItems(3).DisplayValue(12)=   "///////////////////////u9/o8vej7//9Fxe+hrbP/////////////////////////////////"
         Columns(0).ValueItems(3).DisplayValue(13)=   "//////////////9pw+Nq2foviqv/+/b/////////////////////////////////////////////"
         Columns(0).ValueItems(3).DisplayValue(14)=   "///9/v6S0ur29/f///////////////////////////8="
         Columns(0).ValueItems(3).DisplayValue.vt=   9
         Columns(0).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
         Columns(0).ValueItems.Count=   4
         Columns(0).Caption=   "Est"
         Columns(0).DataField=   "tEstado"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Pedido"
         Columns(1).DataField=   "tPedido"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Item"
         Columns(2).DataField=   "tItem"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Fecha"
         Columns(3).DataField=   "tFecha"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).ValueItems(0)._DefaultItem=   0
         Columns(4).ValueItems(0).Value=   ""
         Columns(4).ValueItems(0).Value.vt=   8
         Columns(4).ValueItems(0).DisplayValue=   "0"
         Columns(4).ValueItems(0).DisplayValue.vt=   8
         Columns(4).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(4).ValueItems.Count=   1
         Columns(4).Caption=   "Mensaje"
         Columns(4).DataField=   "tMensaje"
         Columns(4).ButtonPicture.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
         Columns(4).ButtonPicture(0)=   "bHQAADYMAABCTTYMAAAAAAAANgAAACgAAAAgAAAAIAAAAAEAGAAAAAAAAAwAAAAAAAAAAAAAAAAA"
         Columns(4).ButtonPicture(1)=   "AAAAAAD////////////////////////////+/fzjzLHIml7RrXzx59v/////////////////////"
         Columns(4).ButtonPicture(2)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(4).ButtonPicture(3)=   "///////////07OPVq23fqkrEfwbNpnX/////////////////////////////////////////////"
         Columns(4).ButtonPicture(4)=   "///////////////////////////////////////////////////////////////x6N7UqWrgrVDJ"
         Columns(4).ButtonPicture(5)=   "hg3OqXf/////////////////////////////////////////////////////////////////////"
         Columns(4).ButtonPicture(6)=   "///////////////////////////////////////z6uDUqWrfrE/IhQ3PqXn/////////////////"
         Columns(4).ButtonPicture(7)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(4).ButtonPicture(8)=   "///////////////z6uDUqWrfrE7IhQzPqXn/////////////////////////////////////////"
         Columns(4).ButtonPicture(9)=   "///////////////////////////////////////////////////////////////////z6uDUqWnf"
         Columns(4).ButtonPicture(10)=   "rE7IhAzPqXn/////////////////////////////////////////////////////////////////"
         Columns(4).ButtonPicture(11)=   "///////////////////////////////////////////z6uDUqWnfq03IgwvPqXn/////////////"
         Columns(4).ButtonPicture(12)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(4).ButtonPicture(13)=   "///////////////////z6uDUqWnfq03IgwvPqXn/////////////////////////////////////"
         Columns(4).ButtonPicture(14)=   "///////////////////////////////////////////////////////////////////////z6uDU"
         Columns(4).ButtonPicture(15)=   "qWjeq03IgwvPqXn/////////////////////////////////////////////////////////////"
         Columns(4).ButtonPicture(16)=   "///////////////////////////////////////////////17ubRpWLcqEbIhA3Oqnz/////////"
         Columns(4).ButtonPicture(17)=   "///////////////////////////////////////v7+/ExMSTk5OJiYmvr6/f39//////////////"
         Columns(4).ButtonPicture(18)=   "/////////////////////v3bvZfhvHzZoj/OjRnGjz317eX/////////////////////////////"
         Columns(4).ButtonPicture(19)=   "///////////r6ee/q4zHlEPZhwnBjT2ahmZ2dHLJycn///////////////////////////7ZvJjY"
         Columns(4).ButtonPicture(20)=   "qFfryYrXnzrRkiLKiBG+iDn48uz////////////////////////////////49vTRrXXpmRv1qzX+"
         Columns(4).ButtonPicture(21)=   "t0b1qzXpmRuui1J2dHLf39/////////////////////u4tPUoVDqw33mv3nYoDvSlCbOjhzFgArT"
         Columns(4).ButtonPicture(22)=   "sH///v7////////////////////////////r1rfpmRv+uEf/uUn/uUn/uUn+uEfpmRuahmavr6//"
         Columns(4).ButtonPicture(23)=   "///////////////48+7QnlLnu2bowHjkuWvYoTvTlinPkB/LiRK9dgLjzbL/////////////////"
         Columns(4).ButtonPicture(24)=   "///////////irl3zrDr+uEn/uUn/uUn/uUn/uUn1qzXBjDyJiYn////////////z6+HKmFHrvGPn"
         Columns(4).ButtonPicture(25)=   "vGzmumfitF7ZoTvUmSzQkiDLihbIhAi5cgDYvZz////////////////////////ciw36yn3+vlz/"
         Columns(4).ButtonPicture(26)=   "uUr/uUn/uUn/uUn+t0bZhwmTk5P////////7+PTXsXvuwWzoumfis2LerVncqFPWoELTmzjPlCzK"
         Columns(4).ButtonPicture(27)=   "ix/FgQzDegC6dwnr3cz////////////////////lsmH73a78xmz+vlT/uUn/uUn/uUn1qzXGkkHE"
         Columns(4).ButtonPicture(28)=   "xMT////////cwaDSoVnWplrHjzrGjTDFjCzAhim+hC3BiTjDjT3DjTzBizjBhyzAgiK/ij359O//"
         Columns(4).ButtonPicture(29)=   "///////////////037/pmRv62qX8xWz9v1n+uUn+uEfpmRu+qYnv7+/////o2cW7fibFhxTKixDU"
         Columns(4).ButtonPicture(30)=   "lQ3XlgDXlADPigDHfwDHggfNjBjRlCjVnkDUoUrSoVDOoFzQr4X//v3////////////+/PrtyJDp"
         Columns(4).ButtonPicture(31)=   "mRv74bn7z4f0qzbpmRvRrHPq6Ob////y6eC4dAPSjwDXlADZlgDbmQDZmADTkQDMiADGgADFgAfK"
         Columns(4).ButtonPicture(32)=   "iBTPjx7SlCPTlijWmy/aoj3csGTo18L////////////////+/Prw27vlsF/ciw3grFvp1LT39fL/"
         Columns(4).ButtonPicture(33)=   "///////p2sq7dQDZlQDamADbmgDamgDXlQDRjQDLhgDGgADEfgXIhRDMjBnQkCHTlSjUlynUlyna"
         Columns(4).ButtonPicture(34)=   "pELp2MH4+PjS0tKbm5uPj4+7u7vs7Oz////////////////////////////////o18HJlDfMigDQ"
         Columns(4).ButtonPicture(35)=   "jADVkQDVkgDSjgDNhwDIggDGfwDJhArNihLNixfMjB/KjCPSoFHbvZj6+fi00dVRvdMNzvhFsshr"
         Columns(4).ButtonPicture(36)=   "h4yOjo7s7Oz////////////////////////////////38evjz7jkzq/PoljCiSe9fha7exS4dxG5"
         Columns(4).ButtonPicture(37)=   "eBa6ex68gCbGlEnbvpbo2MTw5tv////N6e4J2f0A7P8A+v8A7P8H2Ptrh4y7u7v/////////////"
         Columns(4).ButtonPicture(38)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(4).ButtonPicture(39)=   "//////9v2vAA7P8B/P8A/P8A/P8A7P9FscePj4//////////////////////////////////////"
         Columns(4).ButtonPicture(40)=   "//////////////////////////////////////////////////////////8T1P5o+v8U+/8A/P8A"
         Columns(4).ButtonPicture(41)=   "/P8A+v8Nzvibm5v/////////////////////////////////////////////////////////////"
         Columns(4).ButtonPicture(42)=   "//////////////////////////////////904PWv+f8q+v8M/P8A/P8A7P9Qu9HS0tL/////////"
         Columns(4).ButtonPicture(43)=   "///////////////////////////////////////////////////////////////////w8PC/v7+M"
         Columns(4).ButtonPicture(44)=   "jIypqane3t7R7PAL2/7P+/+i/P8A7P8J2fyzztL4+Pj/////////////////////////////////"
         Columns(4).ButtonPicture(45)=   "///////////////////////////////////////08/ORjb0zMLovLbVhXo53dnfe3t7U7/N14PQS"
         Columns(4).ButtonPicture(46)=   "1P5u2e/M5+v7+/v/////////////////////////////////////////////////////////////"
         Columns(4).ButtonPicture(47)=   "//////////////+uqtoICNoTE/ASEvEICNpiX46pqan/////////////////////////////////"
         Columns(4).ButtonPicture(48)=   "//////////////////////////////////////////////////////////////////9FQ8wSEvEZ"
         Columns(4).ButtonPicture(49)=   "Gf0YF/4SEvEuLbWMjIz/////////////////////////////////////////////////////////"
         Columns(4).ButtonPicture(50)=   "//////////////////////////////////////////9IRs95ePYhIv4ZGf0SEvEyMLm/v7//////"
         Columns(4).ButtonPicture(51)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(4).ButtonPicture(52)=   "//////////////////+6tuQICNp4ePYTE/AICNqQi7rw8PD/////////////////////////////"
         Columns(4).ButtonPicture(53)=   "///////////////////////////////////////////////////////////////////////8+vq5"
         Columns(4).ButtonPicture(54)=   "teRIRc5FQ8ysqNjz8vL///////////////////////////////////////////////////////8="
         Columns(4).ButtonPicture.vt=   9
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=847"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=767"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=257"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2461"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2381"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=260"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(15)=   "Column(3).Width=3678"
         Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=3598"
         Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=260"
         Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(20)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         Appearance      =   2
         DefColWidth     =   0
         HeadLines       =   1.5
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   13160660
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         DirectionAfterEnter=   1
         DirectionAfterTab=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         CollapseColor   =   128
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&H8000000F&"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=47,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=96,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=48,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=49,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=50,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=91,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=94,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=95,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=102,.parent=47,.alignment=2"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=99,.parent=48,.alignment=0"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=100,.parent=49"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=101,.parent=91"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=118,.parent=47"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=115,.parent=48,.alignment=0"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=116,.parent=49"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=117,.parent=91"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=20,.parent=47"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=17,.parent=48"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=18,.parent=49"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=19,.parent=91"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=122,.parent=47"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=119,.parent=48,.alignment=0"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=120,.parent=49"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=121,.parent=91"
         _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=16,.parent=47"
         _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=13,.parent=48"
         _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=14,.parent=49"
         _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=15,.parent=91"
         _StyleDefs(52)  =   "Named:id=33:Normal"
         _StyleDefs(53)  =   ":id=33,.parent=0"
         _StyleDefs(54)  =   "Named:id=34:Heading"
         _StyleDefs(55)  =   ":id=34,.parent=33,.alignment=3,.valignment=2,.bgcolor=&H8000000F&"
         _StyleDefs(56)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1"
         _StyleDefs(57)  =   "Named:id=35:Footing"
         _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(59)  =   "Named:id=36:Selected"
         _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=37:Caption"
         _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(63)  =   "Named:id=38:HighlightRow"
         _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=39:EvenRow"
         _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(67)  =   "Named:id=40:OddRow"
         _StyleDefs(68)  =   ":id=40,.parent=33"
         _StyleDefs(69)  =   "Named:id=41:RecordSelector"
         _StyleDefs(70)  =   ":id=41,.parent=34"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Detalles del Proceso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   210
         Left            =   1245
         TabIndex        =   12
         Top             =   -15
         Width           =   1470
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   11010
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6495
      Width           =   11070
      Begin VB.CommandButton cmdEmite 
         Caption         =   "Reporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   6600
         Picture         =   "frmDescargo.frx":164A
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Reporte Detalle de descargo de Ventas"
         Top             =   60
         Width           =   1155
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   9240
         Picture         =   "frmDescargo.frx":1B7C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Descargar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   7920
         Picture         =   "frmDescargo.frx":1C6E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Descargar"
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Verifique que todos los usuarios se encuentren fuera del sistema de Almacen para poder ejecutar este proceso."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   1290
         TabIndex        =   9
         Top             =   60
         Width           =   4485
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Cuidado :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   15
         TabIndex        =   8
         Top             =   15
         Width           =   1230
      End
   End
   Begin MSComCtl2.Animation aniVideo 
      Height          =   540
      Left            =   135
      TabIndex        =   6
      Top             =   5325
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   953
      _Version        =   393216
      FullWidth       =   49
      FullHeight      =   36
   End
   Begin VB.Frame Frame1 
      Height          =   5805
      Left            =   30
      TabIndex        =   13
      Top             =   480
      Width           =   4410
      Begin MSComctlLib.ProgressBar pbProgress 
         Height          =   330
         Left            =   480
         TabIndex        =   24
         Top             =   2550
         Visible         =   0   'False
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   582
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar pbProgreso 
         Height          =   330
         Left            =   480
         TabIndex        =   22
         Top             =   3570
         Visible         =   0   'False
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   582
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   1140
         TabIndex        =   2
         Top             =   825
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   87359489
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   345
         Left            =   1140
         TabIndex        =   0
         Top             =   345
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   87359489
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   345
         Left            =   2625
         TabIndex        =   1
         Top             =   345
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   87359491
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   345
         Left            =   2625
         TabIndex        =   3
         Top             =   825
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   87359491
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Label1"
         ForeColor       =   &H000040C0&
         Height          =   615
         Left            =   1605
         TabIndex        =   27
         Top             =   4680
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Label1"
         ForeColor       =   &H00008000&
         Height          =   615
         Left            =   3165
         TabIndex        =   26
         Top             =   4530
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblProgress 
         AutoSize        =   -1  'True
         Caption         =   "Procesando Registro Nº 1 de 100"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   480
         TabIndex        =   25
         Top             =   2325
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblProgreso 
         AutoSize        =   -1  'True
         Caption         =   "Procesando Registro Nº 1 de 100"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   480
         TabIndex        =   23
         Top             =   3345
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404000&
         X1              =   0
         X2              =   4590
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   165
         TabIndex        =   21
         Top             =   915
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicial :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   90
         TabIndex        =   20
         Top             =   405
         Width           =   975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Pasos para el descargo de ventas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   210
         Left            =   555
         TabIndex        =   18
         Top             =   0
         Width           =   2490
      End
      Begin VB.Label lblProceso 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Realizando conección"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   0
         Left            =   480
         TabIndex        =   17
         Top             =   1530
         Width           =   1905
      End
      Begin VB.Label lblProceso 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copiando temporales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   1
         Left            =   480
         TabIndex        =   16
         Top             =   2010
         Width           =   1830
      End
      Begin VB.Label lblProceso 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando pedidos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   2
         Left            =   480
         TabIndex        =   15
         Top             =   3030
         Width           =   1770
      End
      Begin VB.Label lblProceso 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Finalizando el proceso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   3
         Left            =   480
         TabIndex        =   14
         Top             =   4140
         Width           =   1950
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   0
         Left            =   120
         Picture         =   "frmDescargo.frx":21F8
         Top             =   1530
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   1
         Left            =   120
         Picture         =   "frmDescargo.frx":240B
         Top             =   2010
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   2
         Left            =   120
         Picture         =   "frmDescargo.frx":261E
         Top             =   3015
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   3
         Left            =   120
         Picture         =   "frmDescargo.frx":2831
         Top             =   4125
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descargo de Ventas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   165
      TabIndex        =   19
      Top             =   90
      Width           =   2160
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   465
      Left            =   0
      Top             =   0
      Width           =   11370
   End
End
Attribute VB_Name = "frmDescargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsTransferencia As Recordset
Dim RsCierre        As Recordset
Dim lCierre         As Boolean
Dim fFechaCierre    As Date
Dim lFechaCierre    As Boolean
Dim clsDescAlmacen  As clsAlmacen

'''''''''''''''''======================
' nombre de store : [sp_TempTrans_LlenaDescargos]  BD ALMACE_NORKYS
Dim sTemporal As String
Dim sPedidosSinDescargo As String
Dim sDtemporal As String
Dim cont As Double

Dim RsTempo As Recordset
Dim RsDtemporal As Recordset
Dim RsProducto As Recordset
Dim RsSubStock As Recordset
Dim RsRecetaVenta As Recordset
Dim RsRecetaVentaDetalle As Recordset

Dim lDescargo As Boolean
Dim lDescargoPedido As Boolean

Dim cCorrect As Double
Dim cError As Double
Public msgError As String

Dim EquipoIp As String
Dim EquipoName As String
Dim EquipoUser As String
Dim DscgFecIni As String
Dim DscgFecFin As String
Dim DscgFecSvrIni As String
Dim DscgFecSvrFin As String
Dim FechaTransaccion As String

Dim dsrReporte As New dsrRepDescargo
Dim fInicio As String
Dim fFinal As String
Dim RsReporte As Recordset


Sub Progreso(Etiqueta As String, cont As Double, Registros As Double)
    lblProgress.Caption = Etiqueta
    pbProgress.value = (cont * 100) / Registros
End Sub

Sub InicializaTablas()
    CnAlmacen.Execute "Delete From TPARAMETRO"
    CnAlmacen.Execute "Delete From TSUBFAMILIA"
    CnAlmacen.Execute "Delete From TTABLA"
    CnAlmacen.Execute "Delete From TPRODUCTO"
    CnAlmacen.Execute "Delete From MRECETAVENTA"
    CnAlmacen.Execute "Delete From DRECETAVENTA"
    CnAlmacen.Execute "Delete From MRECETAPROPIEDAD"
    CnAlmacen.Execute "Delete From DRECETAPROPIEDAD"
    CnAlmacen.Execute "Delete From MRECETABASE"
    CnAlmacen.Execute "Delete From DRECETABASE"
End Sub

Sub MuestraEvento(Indice As Integer)
    lblProceso(Indice).ForeColor = &H80000012
    imgProceso(Indice).Visible = True
End Sub


Sub LlenaDatosInforest()
    'Inserta el Cierre
    lCierre = Calcular("select lCierre as Codigo from tParametro", CnAlmacen)   'era cnalmacen
    fFechaCierre = Calcular("select getdate() as Codigo", Cn)
    If lCierre Then
        If Calcular("select count(fCierre) as Codigo from mCierre where fCierre='" & Format(fFechaCierre, "YYYY/mm/dd") & "'", CnAlmacen) = 0 Then
            Isql = "SpInsmCierre '" & Format(fFechaCierre, "YYYY/mm/dd") & "', '" & sUsuario & "', 0"
            CnAlmacen.Execute Isql
        End If
        
        Isql = "SpLismCierre '" & Format(fFechaCierre, "YYYY/mm/dd") & "'"
        Set RsCierre = Lib.OpenRecordset(Isql, CnAlmacen)
        fFechaCierre = Format(dtpFecFin.value, "YYYY/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm:ss")
    End If
    
    'Desarollador: ELDCQ 05/12/2018
    'Llena la tabla temporal con los datos que se han de descargar
    Cn.Execute " exec usp_inforest_DescargoVenta '" & sAlmacenMDB & "','" & Format(dtpFecIni.value, "yyyymmdd") & " " & Format(dtpHorIni.value, "HH:mm") & "','" & Format(dtpFecFin.value, "yyyymmdd") & " " & Format(dtpHorFin.value, "HH:mm") & "','" & sTemporal & "','" & sLocal & "','',1"
    
'    'Ventas a descargar
'    'Venta por Receta
'    Isql = "SELECT T1.tCodigoPedido, T1.fFecha, T1.tCodigoProducto AS Plato, T1.nCantidad, T1.tItem, T1.tDescargo, T1.tEnlace, T1.tTipoPedido, " & sAlmacenMDB & ".dbo.DRECETAVENTA.tCodigoProducto, " & sAlmacenMDB & ".dbo.DRECETAVENTA.nCantidad AS nRecetaCantidad, " & sAlmacenMDB & ".dbo.DRECETAVENTA.tSubArea, ISNULL(T1.tSubAlmacen,''), " & sAlmacenMDB & ".dbo.TPRODUCTO.lRecetaBase, " & sAlmacenMDB & ".dbo.DRECETAVENTA.lProducto, T1.tCodigoUnicoEtiqueta, T1.tDocumento, T1.fDiaContable " & _
'           "From ( " & _
'           "SELECT dbo.DPEDIDO.tCodigoPedido, (CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha, dbo.DPEDIDO.tCodigoProducto, dbo.DPEDIDO.nCantidad, dbo.DPEDIDO.tItem, dbo.TPRODUCTO.tDescargo, dbo.TPRODUCTO.tEnlace, dbo.MPEDIDO.tTipoPedido , dbo.DPEDIDO.tSubalmacen, ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,'') as tCodigoUnicoEtiqueta, ISNULL(dbo.DPEDIDO.tDocumento,'') as tDocumento, ISNULL(dbo.MPEDIDO.fDiaContable,'') as fDiaContable " & _
'           "FROM dbo.TCANALVENTA INNER JOIN dbo.DPEDIDO ON dbo.TCANALVENTA.tCodigoCanalVenta = dbo.DPEDIDO.tTipoPedido LEFT OUTER JOIN dbo.TPRODUCTO ON dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido " & _
'           "WHERE (((dbo.MPEDIDO.tEstadoPedido = '02' OR dbo.MPEDIDO.tEstadoPedido = '04' OR dbo.MPEDIDO.tEstadoPedido = '05') AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.DPEDIDO.tEstadoItem = 'N') AND (dbo.MPEDIDO.fFecha        >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "') AND (dbo.MPEDIDO.fFecha        <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and dbo.TPRODUCTO.lCombinacion=0) OR " & _
'           "       ((dbo.MPEDIDO.tEstadoPedido = '02' OR dbo.MPEDIDO.tEstadoPedido = '04' OR dbo.MPEDIDO.tEstadoPedido = '05') AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.DPEDIDO.tEstadoItem = 'N') AND (dbo.MPEDIDO.fProgramacion >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "') AND (dbo.MPEDIDO.fProgramacion <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') AND (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND (dbo.TCANALVENTA.lCanalCentralPedidos = 1) and dbo.TPRODUCTO.lCombinacion=0))" & _
'           ") T1 " & _
'           "INNER JOIN " & sAlmacenMDB & ".dbo.MRECETAVENTA ON T1.tEnlace = " & sAlmacenMDB & ".dbo.MRECETAVENTA.tRecetaVenta INNER JOIN " & sAlmacenMDB & ".dbo.DRECETAVENTA ON " & sAlmacenMDB & ".dbo.MRECETAVENTA.tRecetaVenta = " & sAlmacenMDB & ".dbo.DRECETAVENTA.tRecetaVenta AND " & sAlmacenMDB & ".dbo.MRECETAVENTA.tLocal = " & sAlmacenMDB & ".dbo.DRECETAVENTA.tLocal " & _
'           "INNER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO ON " & sAlmacenMDB & ".dbo.TPRODUCTO.tCodigoProducto = " & sAlmacenMDB & ".dbo.DRECETAVENTA.tCodigoProducto " & _
'           "WHERE (" & sAlmacenMDB & ".dbo.DRECETAVENTA.lDescargo = 1) and ((T1.tTipoPedido = '01' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal1 = 1) or (T1.tTipoPedido = '02' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal2= 1) or (T1.tTipoPedido = '03' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal3= 1) or (T1.tTipoPedido = '04' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal4 = 1) or (T1.tTipoPedido = '05' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal5 = 1)) and " & sAlmacenMDB & ".dbo.MRECETAVENTA.tLocal='" & sLocal & "' And " & sAlmacenMDB & ".dbo.MRECETAVENTA.lActivo = 1"
'    Debug.Print Isql
'    Cn.Execute "insert into " & sTemporal & " " & Isql
'
'
'
'    'Venta por Descargo Directo
'    Isql = "SELECT dbo.DPEDIDO.tCodigoPedido, (CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha, dbo.DPEDIDO.tCodigoProducto AS Plato, dbo.DPEDIDO.nCantidad, dbo.DPEDIDO.tItem, dbo.TPRODUCTO.tDescargo, dbo.TPRODUCTO.tEnlace, dbo.MPEDIDO.tTipoPedido, dbo.TPRODUCTO.tEnlace AS tCodigoProducto, 1 AS nRecetaCantidad, dbo.vArea.tValor AS tSubArea, ISNULL(dbo.DPEDIDO.tSubalmacen,''), 0, 0, ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,''), ISNULL(dbo.DPEDIDO.tDocumento,'') as tDocumento, ISNULL(dbo.MPEDIDO.fDiaContable,'') as fDiaContable " & _
'           "FROM dbo.TCANALVENTA INNER JOIN dbo.DPEDIDO ON dbo.TCANALVENTA.tCodigoCanalVenta = dbo.DPEDIDO.tTipoPedido LEFT OUTER JOIN dbo.TPRODUCTO LEFT OUTER JOIN dbo.vArea ON dbo.TPRODUCTO.tArea = dbo.vArea.Codigo ON dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido " & _
'           "WHERE " & _
'           "(((dbo.MPEDIDO.tEstadoPedido = '02' OR dbo.MPEDIDO.tEstadoPedido = '04' OR dbo.MPEDIDO.tEstadoPedido = '05') AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.DPEDIDO.tEstadoItem = 'N') AND (dbo.TPRODUCTO.tDescargo = 'D') AND (ISNULL(dbo.TPRODUCTO.tEnlace, N'') <> '') AND (dbo.MPEDIDO.fFecha        >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "') AND (dbo.MPEDIDO.fFecha        <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and dbo.TPRODUCTO.lCombinacion=0) OR " & _
'           " ((dbo.MPEDIDO.tEstadoPedido = '02' OR dbo.MPEDIDO.tEstadoPedido = '04' OR dbo.MPEDIDO.tEstadoPedido = '05') AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.DPEDIDO.tEstadoItem = 'N') AND (dbo.TPRODUCTO.tDescargo = 'D') AND (ISNULL(dbo.TPRODUCTO.tEnlace, N'') <> '') AND (dbo.MPEDIDO.fProgramacion >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "') AND (dbo.MPEDIDO.fProgramacion <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') AND (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND (dbo.TCANALVENTA.lCanalCentralPedidos = 1) and dbo.TPRODUCTO.lCombinacion=0)) "
'    Debug.Print Isql
'    Cn.Execute "insert into " & sTemporal & " " & Isql
'
'    'Quita los Sin de las Ventas
'    Cn.Execute "delete from " & sTemporal & " where tCodigoPedido + tItem + Plato + tCodigoProducto in " & _
'               " (SELECT     dbo.TPRODUCTOPROPIEDAD.tCodigoPedido + dbo.TPRODUCTOPROPIEDAD.tItem + dbo.TPRODUCTOPROPIEDAD.tProducto + dbo.TPRODUCTOPROPIEDAD.tEnlace " & _
'               "FROM dbo.TPRODUCTOPROPIEDAD INNER JOIN dbo.MPEDIDO ON dbo.TPRODUCTOPROPIEDAD.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido INNER JOIN dbo.DPEDIDO ON dbo.TPRODUCTOPROPIEDAD.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido AND dbo.TPRODUCTOPROPIEDAD.tItem = dbo.DPEDIDO.tItem INNER JOIN dbo.TCANALVENTA ON dbo.DPEDIDO.tTipoPedido = dbo.TCANALVENTA.tCodigoCanalVenta " & _
'               "WHERE " & _
'               "(((dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = '9999') AND (dbo.MPEDIDO.tEstadoPedido = '02' OR dbo.MPEDIDO.tEstadoPedido = '04' OR dbo.MPEDIDO.tEstadoPedido = '05') AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.DPEDIDO.tEstadoItem = 'N') AND (dbo.MPEDIDO.fFecha        >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "') AND (dbo.MPEDIDO.fFecha        <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0)) OR " & _
'               " ((dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = '9999') AND (dbo.MPEDIDO.tEstadoPedido = '02' OR dbo.MPEDIDO.tEstadoPedido = '04' OR dbo.MPEDIDO.tEstadoPedido = '05') AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.DPEDIDO.tEstadoItem = 'N') AND (dbo.MPEDIDO.fProgramacion >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "') AND (dbo.MPEDIDO.fProgramacion <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') AND (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND (dbo.TCANALVENTA.lCanalCentralPedidos = 1)))) "
'
'    'Propiedades con Receta
'        Isql = "SELECT T1.tCodigoPedido, T1.fFecha, T1.Plato, T1.nCantidad, T1.tItem, T1.tDescargo, T1.tEnlace, T1.tTipoPedido, " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.tCodigoProducto, " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.nCantidad AS nRecetaCantidad, " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.tSubArea, ISNULL(T1.tSubAlmacen,''), " & sAlmacenMDB & ".dbo.tProducto.lRecetaBase , " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.lProducto, T1.tCodigoUnicoEtiqueta, T1.tDocumento, T1.fDiaContable " & _
'            "From (SELECT dbo.TPRODUCTOPROPIEDAD.tCodigoPedido,(CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha, dbo.TPRODUCTOPROPIEDAD.tProducto AS Plato, dbo.DPEDIDO.nCantidad * isnull(tproductopropiedad.ncantidad,1) ncantidad, dbo.DPEDIDO.tItem, (CASE WHEN LEN(dbo.TPRODUCTOPROPIEDAD.tEnlace)= 5 THEN 'R' ELSE 'D' END) AS tDescargo, dbo.TPRODUCTOPROPIEDAD.tEnlace, dbo.MPEDIDO.tTipoPedido, dbo.DPEDIDO.tSubalmacen, ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,'') as tCodigoUnicoEtiqueta, ISNULL(dbo.DPEDIDO.tDocumento,'') as tDocumento, ISNULL(dbo.MPEDIDO.fDiaContable,'') as fDiaContable " & _
'            "FROM dbo.DPEDIDO INNER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido INNER JOIN dbo.TPRODUCTOPROPIEDAD ON dbo.DPEDIDO.tCodigoPedido = dbo.TPRODUCTOPROPIEDAD.tCodigoPedido AND dbo.DPEDIDO.tItem = dbo.TPRODUCTOPROPIEDAD.tItem INNER JOIN dbo.TPRODUCTO ON dbo.TPRODUCTOPROPIEDAD.tProducto = dbo.TPRODUCTO.tCodigoProducto " & _
'            "INNER JOIN dbo.TCANALVENTA ON dbo.DPEDIDO.tTipoPedido = dbo.TCANALVENTA.tCodigoCanalVenta WHERE " & _
'            "(((dbo.MPEDIDO.tEstadoPedido = '02' OR dbo.MPEDIDO.tEstadoPedido = '04' OR dbo.MPEDIDO.tEstadoPedido = '05') AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.MPEDIDO.fFecha        >='" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "') AND (dbo.MPEDIDO.fFecha        <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and dbo.TPRODUCTO.lCombinacion=0) OR " & _
'            " ((dbo.MPEDIDO.tEstadoPedido = '02' OR dbo.MPEDIDO.tEstadoPedido = '04' OR dbo.MPEDIDO.tEstadoPedido = '05') AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.MPEDIDO.fProgramacion >='" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "') AND (dbo.MPEDIDO.fProgramacion <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') AND (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND (dbo.TCANALVENTA.lCanalCentralPedidos = 1) and dbo.TPRODUCTO.lCombinacion=0))" & _
'            ") T1 " & _
'            "INNER JOIN  " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD ON T1.tEnlace = " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD.tRecetaPropiedad INNER JOIN " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD ON " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD.tRecetaPropiedad = " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.tRecetaPropiedad AND " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD.tLocal = " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.tLocal " & _
'            "INNER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO ON " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.tCodigoProducto = " & sAlmacenMDB & ".dbo.TPRODUCTO.tCodigoProducto " & _
'            "WHERE (" & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.lDescargo = 1) and " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD.tLocal='" & sLocal & "' "
'    Debug.Print Isql
'    Cn.Execute "insert into " & sTemporal & " " & Isql
'
'    'Propiedades con Descargo Directo
'    Isql = "SELECT dbo.TPRODUCTOPROPIEDAD.tCodigoPedido,(CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha, dbo.DPEDIDO.tCodigoProducto AS Plato, dbo.DPEDIDO.nCantidad * isnull(tproductopropiedad.ncantidad,1) NCANTIDAD, dbo.TPRODUCTOPROPIEDAD.tItem, 'D' AS tDescargo, dbo.TPRODUCTOPROPIEDAD.tEnlace, dbo.MPEDIDO.tTipoPedido, dbo.TPRODUCTOPROPIEDAD.tEnlace AS tCodigoProducto, 1 AS nRecetaCantidad, dbo.vArea.tValor AS tSubArea,ISNULL(dbo.DPEDIDO.tSubalmacen, N''), 0, 0, ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,''), ISNULL(dbo.DPEDIDO.tDocumento,'') as tDocumento, ISNULL(dbo.MPEDIDO.fDiaContable,'') as fDiaContable " & _
'           "FROM dbo.DPEDIDO INNER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido INNER JOIN dbo.TPROPIEDAD INNER JOIN dbo.TPRODUCTOPROPIEDAD ON dbo.TPROPIEDAD.tCodigoPropiedad = dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad AND dbo.TPROPIEDAD.tProducto = dbo.TPRODUCTOPROPIEDAD.tProducto ON dbo.DPEDIDO.tCodigoPedido = dbo.TPRODUCTOPROPIEDAD.tCodigoPedido AND dbo.DPEDIDO.tItem = dbo.TPRODUCTOPROPIEDAD.tItem INNER JOIN dbo.vArea ON dbo.TPROPIEDAD.tArea = dbo.vArea.Codigo INNER JOIN dbo.TCANALVENTA ON dbo.DPEDIDO.tTipoPedido = dbo.TCANALVENTA.tCodigoCanalVenta inner join dbo.TPRODUCTO on dbo.TPRODUCTO.tCodigoProducto=dbo.DPEDIDO.tCodigoProducto " & _
'           "WHERE " & _
'           "(((dbo.MPEDIDO.tEstadoPedido = '02' OR dbo.MPEDIDO.tEstadoPedido = '04' OR dbo.MPEDIDO.tEstadoPedido = '05') AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (LEN(dbo.TPROPIEDAD.tEnlace) = 7) AND (dbo.MPEDIDO.fFecha        >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "') AND (dbo.MPEDIDO.fFecha        <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and dbo.TPRODUCTO.lCombinacion=0) OR " & _
'           " ((dbo.MPEDIDO.tEstadoPedido = '02' OR dbo.MPEDIDO.tEstadoPedido = '04' OR dbo.MPEDIDO.tEstadoPedido = '05') AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (LEN(dbo.TPROPIEDAD.tEnlace) = 7) AND (dbo.MPEDIDO.fProgramacion >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "') AND (dbo.MPEDIDO.fProgramacion <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') AND (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND (dbo.TCANALVENTA.lCanalCentralPedidos = 1) and dbo.TPRODUCTO.lCombinacion=0)) "
'    Debug.Print Isql
'    Cn.Execute "insert into " & sTemporal & " " & Isql
'
'    'Combos por Recetas
'    Isql = "SELECT T1.tCodigoPedido, T1.fFecha, T1.tCodigoProducto AS Plato, T1.nCantidad, T1.tItem, T1.tDescargo, T1.tEnlace, T1.tTipoPedido, " & sAlmacenMDB & ".dbo.DRECETAVENTA.tCodigoProducto, " & sAlmacenMDB & ".dbo.DRECETAVENTA.nCantidad AS nRecetaCantidad, " & sAlmacenMDB & ".dbo.DRECETAVENTA.tSubArea, ISNULL(T1.tSubAlmacen,''), " & sAlmacenMDB & ".dbo.TPRODUCTO.lRecetaBase, " & sAlmacenMDB & ".dbo.DRECETAVENTA.lProducto,T1.tCodigoUnicoEtiqueta, T1.tDocumento, T1.fDiaContable " & _
'           "From (SELECT dbo.DPEDIDO.tCodigoPedido, (CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha, dbo.CPEDIDO.tProductoCombo AS tCodigoProducto, dbo.CPEDIDO.nCantidad, dbo.DPEDIDO.tItem, TPRODUCTO_1.tDescargo, TPRODUCTO_1.tEnlace, dbo.MPEDIDO.tTipoPedido , dbo.DPEDIDO.tSubalmacen, ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,'') as tCodigoUnicoEtiqueta, ISNULL(dbo.DPEDIDO.tDocumento,'') as tDocumento, ISNULL(dbo.MPEDIDO.fDiaContable,'') as fDiaContable " & _
'           "FROM dbo.CPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.CPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido AND dbo.CPEDIDO.tItem = dbo.DPEDIDO.tItem INNER JOIN dbo.TPRODUCTO ON dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto INNER JOIN dbo.TPRODUCTO AS TPRODUCTO_1 ON dbo.CPEDIDO.tProductoCombo = TPRODUCTO_1.tCodigoProducto INNER JOIN dbo.TCANALVENTA ON dbo.DPEDIDO.tTipoPedido = dbo.TCANALVENTA.tCodigoCanalVenta RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido " & _
'           "WHERE " & _
'           "(((dbo.MPEDIDO.tEstadoPedido = '02' OR dbo.MPEDIDO.tEstadoPedido = '04' OR dbo.MPEDIDO.tEstadoPedido = '05') AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.DPEDIDO.tEstadoItem = 'N') AND (dbo.MPEDIDO.fFecha        >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "') AND (dbo.MPEDIDO.fFecha        <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and dbo.TPRODUCTO.lCombinacion=1) OR " & _
'           " ((dbo.MPEDIDO.tEstadoPedido = '02' OR dbo.MPEDIDO.tEstadoPedido = '04' OR dbo.MPEDIDO.tEstadoPedido = '05') AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.DPEDIDO.tEstadoItem = 'N') AND (dbo.MPEDIDO.fProgramacion >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "') AND (dbo.MPEDIDO.fProgramacion <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') AND (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND (dbo.TCANALVENTA.lCanalCentralPedidos = 1) and dbo.TPRODUCTO.lCombinacion=1))" & _
'           ") T1 " & _
'           "INNER JOIN " & sAlmacenMDB & ".dbo.MRECETAVENTA ON T1.tEnlace = " & sAlmacenMDB & ".dbo.MRECETAVENTA.tRecetaVenta INNER JOIN " & sAlmacenMDB & ".dbo.DRECETAVENTA ON " & sAlmacenMDB & ".dbo.MRECETAVENTA.tRecetaVenta = " & sAlmacenMDB & ".dbo.DRECETAVENTA.tRecetaVenta AND " & sAlmacenMDB & ".dbo.MRECETAVENTA.tLocal = " & sAlmacenMDB & ".dbo.DRECETAVENTA.tLocal INNER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO ON " & sAlmacenMDB & ".dbo.TPRODUCTO.tCodigoProducto = " & sAlmacenMDB & ".dbo.DRECETAVENTA.tCodigoProducto " & _
'           "WHERE (" & sAlmacenMDB & ".dbo.DRECETAVENTA.lDescargo = 1) and ((T1.tTipoPedido = '01' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal1 = 1) or (T1.tTipoPedido = '02' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal2= 1) or (T1.tTipoPedido = '03' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal3= 1) or (T1.tTipoPedido = '04' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal4 = 1) or (T1.tTipoPedido = '05' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal5 = 1)) and " & sAlmacenMDB & ".dbo.MRECETAVENTA.tLocal='" & sLocal & "' "
'    Debug.Print Isql
'    Cn.Execute "insert into " & sTemporal & " " & Isql
'
'    'Combo por Descargo Directo
'    Isql = "SELECT dbo.DPEDIDO.tCodigoPedido, (CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha, dbo.CPEDIDO.tProductoCombo AS Plato, dbo.CPEDIDO.nCantidad, dbo.DPEDIDO.tItem, TPRODUCTO_1.tDescargo, TPRODUCTO_1.tEnlace, dbo.MPEDIDO.tTipoPedido, TPRODUCTO_1.tEnlace AS tCodigoProducto, 1 AS nRecetaCantidad, dbo.vArea.tValor AS tSubArea, ISNULL(dbo.DPEDIDO.tSubalmacen, N'') , 0 , 0, ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,''), ISNULL(dbo.DPEDIDO.tDocumento,'') as tDocumento, ISNULL(dbo.MPEDIDO.fDiaContable,'') as fDiaContable " & _
'           "FROM dbo.CPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.CPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido AND dbo.CPEDIDO.tItem = dbo.DPEDIDO.tItem INNER JOIN dbo.TPRODUCTO ON dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto INNER JOIN dbo.TPRODUCTO AS TPRODUCTO_1 ON dbo.CPEDIDO.tProductoCombo = TPRODUCTO_1.tCodigoProducto INNER JOIN dbo.vArea ON TPRODUCTO_1.tArea = dbo.vArea.Codigo INNER JOIN dbo.TCANALVENTA ON dbo.DPEDIDO.tTipoPedido = dbo.TCANALVENTA.tCodigoCanalVenta RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido " & _
'           "WHERE " & _
'           "(((dbo.MPEDIDO.tEstadoPedido = '02' OR dbo.MPEDIDO.tEstadoPedido = '04' OR dbo.MPEDIDO.tEstadoPedido = '05') AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.DPEDIDO.tEstadoItem = 'N') AND (TPRODUCTO_1.tDescargo = 'D') AND (dbo.MPEDIDO.fFecha        >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "') AND (dbo.MPEDIDO.fFecha        <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and dbo.TPRODUCTO.lCombinacion=1) OR " & _
'           " ((dbo.MPEDIDO.tEstadoPedido = '02' OR dbo.MPEDIDO.tEstadoPedido = '04' OR dbo.MPEDIDO.tEstadoPedido = '05') AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.DPEDIDO.tEstadoItem = 'N') AND (TPRODUCTO_1.tDescargo = 'D') AND (dbo.MPEDIDO.fProgramacion >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "') AND (dbo.MPEDIDO.fProgramacion <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') AND (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND (dbo.TCANALVENTA.lCanalCentralPedidos = 1) and dbo.TPRODUCTO.lCombinacion=1)) "
'    Debug.Print Isql
'    Cn.Execute "insert into " & sTemporal & " " & Isql
'
'    'Quita los Sin de los Combos
'    Cn.Execute "delete from " & sTemporal & " where tcodigoPedido + tItem + Plato + tCodigoProducto " & _
'               "in (SELECT     dbo.TCOMBOPROPIEDAD.tCodigoPedido + dbo.TCOMBOPROPIEDAD.tItem + dbo.TCOMBOPROPIEDAD.tProducto + dbo.TCOMBOPROPIEDAD.tEnlace FROM dbo.TCOMBOPROPIEDAD INNER JOIN dbo.MPEDIDO ON dbo.TCOMBOPROPIEDAD.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido CROSS JOIN dbo.TCANALVENTA " & _
'               "WHERE " & _
'               "(((dbo.MPEDIDO.tEstadoPedido = '02' OR dbo.MPEDIDO.tEstadoPedido = '04' OR dbo.MPEDIDO.tEstadoPedido = '05') AND (dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = '9999') AND (dbo.MPEDIDO.fFecha        >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "') AND (dbo.MPEDIDO.fFecha        <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0)) OR " & _
'               " ((dbo.MPEDIDO.tEstadoPedido = '02' OR dbo.MPEDIDO.tEstadoPedido = '04' OR dbo.MPEDIDO.tEstadoPedido = '05') AND (dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = '9999') AND (dbo.MPEDIDO.fProgramacion >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "') AND (dbo.MPEDIDO.fProgramacion <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') AND (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND (dbo.TCANALVENTA.lCanalCentralPedidos = 1))))"
'
'    'Propiedades de los Combos con Recetas
'    Isql = "SELECT T1.tCodigoPedido, T1.fFecha, T1.Plato, T1.nCantidad, T1.tItem, T1.tDescargo, T1.tEnlace, T1.tTipoPedido, " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.tCodigoProducto, " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.nCantidad AS nRecetaCantidad, " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.tSubArea, ISNULL(T1.tSubAlmacen,''), 0, 0, T1.tCodigoUnicoEtiqueta, T1.tDocumento, T1.fDiaContable " & _
'           "FROM (SELECT     dbo.TCOMBOPROPIEDAD.tCodigoPedido, (CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha, dbo.CPEDIDO.tProductoCombo AS Plato, dbo.CPEDIDO.nCantidad * isnull(tcombopropiedad.ncantidad,1) ncantidad, dbo.TCOMBOPROPIEDAD.tItem, (CASE WHEN LEN(dbo.TCOMBOPROPIEDAD.tEnlace) = 5 THEN 'R' ELSE 'D' END) AS tDescargo, dbo.TCOMBOPROPIEDAD.tEnlace, dbo.MPEDIDO.tTipoPedido, dbo.DPEDIDO.tSubalmacen, ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,'') as  tCodigoUnicoEtiqueta, ISNULL(dbo.DPEDIDO.tDocumento,'') as tDocumento, ISNULL(dbo.MPEDIDO.fDiaContable,'') as fDiaContable " & _
'           "FROM dbo.TPRODUCTO INNER JOIN dbo.CPEDIDO INNER JOIN dbo.TCOMBOPROPIEDAD ON dbo.CPEDIDO.tCodigoPedido = dbo.TCOMBOPROPIEDAD.tCodigoPedido AND dbo.CPEDIDO.tItem = dbo.TCOMBOPROPIEDAD.tItem AND dbo.CPEDIDO.tItemCombo = dbo.TCOMBOPROPIEDAD.tItemCombo INNER JOIN dbo.DPEDIDO INNER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido ON dbo.CPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido AND dbo.CPEDIDO.tItem = dbo.DPEDIDO.tItem ON dbo.TPRODUCTO.tCodigoProducto = dbo.DPEDIDO.tCodigoProducto INNER JOIN dbo.TCANALVENTA ON dbo.DPEDIDO.tTipoPedido = dbo.TCANALVENTA.tCodigoCanalVenta" & _
'           " WHERE " & _
'           "(((dbo.MPEDIDO.tEstadoPedido = '02' or dbo.MPEDIDO.tEstadoPedido = '04' or dbo.MPEDIDO.tEstadoPedido = '05') AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.MPEDIDO.fFecha        >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "') AND (dbo.MPEDIDO.fFecha        <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and dbo.TPRODUCTO.lCombinacion=1) OR " & _
'           " ((dbo.MPEDIDO.tEstadoPedido = '02' or dbo.MPEDIDO.tEstadoPedido = '04' or dbo.MPEDIDO.tEstadoPedido = '05') AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.MPEDIDO.fProgramacion >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "') AND (dbo.MPEDIDO.fProgramacion <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') AND (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND (dbo.TCANALVENTA.lCanalCentralPedidos = 1) and dbo.TPRODUCTO.lCombinacion=1))" & _
'           ") T1 " & _
'           "INNER JOIN " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD ON T1.tEnlace = " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD.tRecetaPropiedad INNER JOIN " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD ON " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD.tRecetaPropiedad = " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.tRecetaPropiedad AND " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD.tLocal = " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.tLocal " & _
'           " WHERE (" & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.lDescargo = 1) and " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD.tLocal='" & sLocal & "'"
'    Debug.Print Isql
'
'    Cn.Execute "insert into " & sTemporal & " " & Isql
'
'    'Propiedades de los Combos con Descargo Directo
'    Isql = "SELECT dbo.TCOMBOPROPIEDAD.tCodigoPedido, (CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha, dbo.CPEDIDO.tProductoCombo AS Plato, dbo.CPEDIDO.nCantidad * isnull(TCOMBOPROPIEDAD.ncantidad,1) ncantidad, dbo.TCOMBOPROPIEDAD.tItem, 'D' AS tDescargo, dbo.TCOMBOPROPIEDAD.tEnlace, dbo.MPEDIDO.tTipoPedido, dbo.TCOMBOPROPIEDAD.tEnlace AS tCodigoProducto, 1 AS nRecetaCantidad, dbo.TPROPIEDAD.tArea AS tSubArea, ISNULL(dbo.DPEDIDO.tSubalmacen, N''), 0, 0 ,ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,''), ISNULL(dbo.DPEDIDO.tDocumento,'') as tDocumento, ISNULL(dbo.MPEDIDO.fDiaContable,'') as fDiaContable " & _
'           "FROM dbo.CPEDIDO INNER JOIN dbo.TCOMBOPROPIEDAD ON dbo.CPEDIDO.tCodigoPedido = dbo.TCOMBOPROPIEDAD.tCodigoPedido AND dbo.CPEDIDO.tItem = dbo.TCOMBOPROPIEDAD.tItem AND dbo.CPEDIDO.tItemCombo = dbo.TCOMBOPROPIEDAD.tItemCombo INNER JOIN dbo.TPROPIEDAD ON dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND dbo.TCOMBOPROPIEDAD.tProducto = dbo.TPROPIEDAD.tProducto INNER JOIN dbo.DPEDIDO ON dbo.CPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido AND dbo.CPEDIDO.tItem = dbo.DPEDIDO.tItem INNER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido INNER JOIN dbo.TCANALVENTA ON dbo.DPEDIDO.tTipoPedido = dbo.TCANALVENTA.tCodigoCanalVenta INNER JOIN dbo.TPRODUCTO ON dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto " & _
'           "WHERE " & _
'           "(((dbo.MPEDIDO.tEstadoPedido = '02' OR dbo.MPEDIDO.tEstadoPedido = '04' OR dbo.MPEDIDO.tEstadoPedido = '05') AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (LEN(dbo.TPROPIEDAD.tEnlace) = 7) AND (dbo.MPEDIDO.fFecha        >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "') AND (dbo.MPEDIDO.fFecha        <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and dbo.TPRODUCTO.lCombinacion=1) OR " & _
'           " ((dbo.MPEDIDO.tEstadoPedido = '02' OR dbo.MPEDIDO.tEstadoPedido = '04' OR dbo.MPEDIDO.tEstadoPedido = '05') AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (LEN(dbo.TPROPIEDAD.tEnlace) = 7) AND (dbo.MPEDIDO.fProgramacion >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "') AND (dbo.MPEDIDO.fProgramacion <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') AND (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND (dbo.TCANALVENTA.lCanalCentralPedidos = 1) and dbo.TPRODUCTO.lCombinacion=1))"
'    Debug.Print Isql
'    Cn.Execute "insert into " & sTemporal & " " & Isql
End Sub
Sub TraeTemporal()
    Isql = "SELECT tEstado, tPedido, tItem, tFecha, tMensaje FROM " & sDtemporal & " GROUP BY tEstado, tPedido, tItem , tFecha, tMensaje"
    Set RsDtemporal = Lib.OpenRecordset(Isql, Cn)
    Set grdGrilla.DataSource = RsDtemporal
    
    Dim Correcto As New TrueOleDBGrid80.Style
    Correcto.backColor = &HC0FFC0
    'Correcto.Font.Bold = True
    Correcto.ForeColor = &H8000&
    
    Dim Aviso As New TrueOleDBGrid80.Style
    Aviso.backColor = &HC0E0FF
    'Aviso.Font.Bold = True
    Aviso.ForeColor = &H40C0&
    
    grdGrilla.Columns(4).AddRegexCellStyle 0, Correcto, "correctamente"
    grdGrilla.Columns(4).AddRegexCellStyle 0, Aviso, "cierre"
    grdGrilla.Columns(4).Width = 5000
    grdGrilla.Columns(2).Width = 500
    grdGrilla.Columns(3).Width = 1500
End Sub

Sub Inicializa()
    lblProceso(0).ForeColor = &H808080
    lblProceso(1).ForeColor = &H808080
    lblProceso(2).ForeColor = &H808080
    lblProceso(3).ForeColor = &H808080
    imgProceso(0).Visible = False
    imgProceso(1).Visible = False
    imgProceso(2).Visible = False
    imgProceso(3).Visible = False
    
    Isql = "Delete From " & sDtemporal
    Cn.Execute Isql
    TraeTemporal
    
    pbProgreso.value = 0
    pbProgreso.Visible = False
    lblProgreso.Visible = False
    
    aniVideo.AutoPlay = True
    aniVideo.Visible = True
    DoEvents
End Sub

Private Sub cmdEmite_Click()
On Error GoTo fin
     Genera
     If RsReporte.EOF = True Then
       MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
    frmEmite.CRViewer.DisplayGroupTree = False
    dsrReporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    dsrReporte.PaperOrientation = crLandscape
    frmEmite.CRViewer.ViewReport
    frmEmite.Show vbModal
Exit Sub
fin:

End Sub
Private Sub Genera()
On Error GoTo fin

    Dim oComando As clsComando

    
    fInicio = Format(dtpFecIni.value, "yyyyMMdd") & " " & Format(Me.dtpFecIni.value, "HH:mm")
    fFinal = Format(dtpFecFin.value, "yyyyMMdd") & " " & Format(Me.dtpHorFin.value, "HH:mm")
    
    Screen.MousePointer = vbHourglass
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("usp_RepInforest_DescargoVenta", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    oComando.CreateParameter "@Almacen", adVarChar, adParamInput, 50, sAlmacenMDB
    oComando.CreateParameter "@FechaIni", adVarChar, adParamInput, 50, fInicio
    oComando.CreateParameter "@FechaFin", adVarChar, adParamInput, 50, fFinal
    oComando.CreateParameter "@sTemporal", adVarChar, adParamInput, 50, ""
    oComando.CreateParameter "@Local", adVarChar, adParamInput, 50, sLocal
    oComando.CreateParameter "@Grupo", adVarChar, adParamInput, 50, ""
    oComando.CreateParameter "@SubGrupo", adVarChar, adParamInput, 50, ""
    oComando.CreateParameter "@Insumo", adVarChar, adParamInput, 50, ""
    oComando.CreateParameter "@Area", adVarChar, adParamInput, 50, ""
    oComando.CreateParameter "@Descargo", adVarChar, adParamInput, 50, ""
    oComando.CreateParameter "@tipooper", adInteger, adParamInput, 10, 1
    
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set RsReporte = oComando.GetSP()
    'RsReporte.Filter = sCriterio
    
    dsrReporte.DiscardSavedData
    dsrReporte.Database.SetDataSource RsReporte
    'CrtDetalleC.ReportTitle = ""
    dsrReporte.Text8.SetText sRazonSocial
    'Reporte.Text5.SetText localConectado
    frmEmite.CRViewer.ReportSource = dsrReporte

'    If OptDetalle.value = True Then
'         Reporte.DiscardSavedData
'         Reporte.Database.SetDataSource RsReporte
'         'CrtDetalleC.ReportTitle = ""
'         Reporte.Text13.SetText sRazonSocial
'         Reporte.Text5.SetText localConectado
'         frmEmite.CRViewer.ReportSource = Reporte
'    Else
'         CrtResumenC.DiscardSavedData
'         CrtResumenC.Database.SetDataSource RsReporte
'         CrtResumenC.ReportTitle = sTitulo
'         CrtResumenC.Text13.SetText sRazonSocial
'         CrtResumenC.Text3.SetText localConectado
'         frmEmite.CRViewer.ReportSource = CrtResumenC
'    End If

    Screen.MousePointer = vbDefault
Exit Sub
fin:
End Sub

Private Sub cmdOpcion_Click()
    
    Dim baseDatos As String
    Dim servidor As String
    Dim nCant As Double
    
    If dtpFecIni.value > dtpFecFin.value Then MsgBox "Rango de fecha no valido!!!", vbCritical, sMensaje: dtpFecIni.SetFocus: Exit Sub
    
    sTemporal = dbTemporal(sUsuario, 17, "tCodigoPedido", "nVarChar(10)", _
                                         "fFecha", "smalldatetime", _
                                         "Plato", "nVarChar(7)", _
                                         "nCantidad", "Float", _
                                         "tItem", "nVarChar(3)", _
                                         "tDescargo", "nVarChar(1)", _
                                         "tEnlace", "nVarChar(7)", _
                                         "tTipoPedido", "nVarChar(2)", _
                                         "tCodigoProducto", "nVarChar(7)", _
                                         "nRecetaCantidad", "Float", _
                                         "tSubAreaAlm", "nVarChar(3)", _
                                         "tSubAreaInf", "nVarChar(3)", _
                                         "lRecetaBase", "bit", _
                                         "lProducto", "bit", _
                                         "tCodigoUnicoEtiqueta", "nVarChar(26)", _
                                         "tDocumento", "nVarChar(30)", _
                                         "fDiaContable", "smalldatetime")
    
    sPedidosSinDescargo = dbTemporal(sUsuario, 1, "tCodigoPedido", "nVarChar(10)")
                                         
    If MsgBox("Seguro de realizar el proceso???", vbYesNo + vbQuestion, sMensaje) = vbNo Then
        Exit Sub
    End If
    
    grdGrilla.DataSource = Nothing
    Cn.Execute "Delete From " & sDtemporal
    
    If verificaAlmacenRemoto = "ON" Then
        '********************************PROCESO DE LLENADO DE TABLAS********************************
        Screen.MousePointer = vbHourglass
    
        MuestraEvento 0
        
        If lAlmacenRemoto = True Then
            Set CnAlmacenRemoto = New Connection
            CnAlmacenRemoto.Provider = "SQLOLEDB"
            CnAlmacenRemoto.CursorLocation = adUseServer
            CnAlmacenRemoto.ConnectionString = "User ID=" & sUserName & _
                                               ";password=" & sUserPassword & _
                                               ";Data Source=" & sRutaAlmacenRemoto & _
                                               ";Initial Catalog=" & sMDBAlmacenRemoto
            CnAlmacenRemoto.CommandTimeout = 250
            CnAlmacenRemoto.Open
        End If
    
        nCant = 20
        lblProgress.Visible = True
        pbProgress.Visible = True
        pbProgress.value = 0
        
        MuestraEvento 1
        
        InicializaTablas
        
        Call Progreso("Generando parametros generales...", 1, nCant): CreaArchivos "TPARAMETRO"
        Call Progreso("Generando archivo subFamilias...", 2, nCant): CreaArchivos "TSUBFAMILIA"
        Call Progreso("Generando archivo tablas...", 3, nCant): CreaArchivos "TTABLA"
        Call Progreso("Generando archivo productos...", 4, nCant): CreaArchivos "TPRODUCTO"
        Call Progreso("Generando archivo receta venta...", 5, nCant): CreaArchivos "MRECETAVENTA"
        Call Progreso("Generando archivo receta venta detalle...", 6, nCant): CreaArchivos "DRECETAVENTA"
        Call Progreso("Generando archivo receta propiedad...", 7, nCant): CreaArchivos "MRECETAPROPIEDAD"
        Call Progreso("Generando archivo receta propiedad detalle...", 8, nCant): CreaArchivos "DRECETAPROPIEDAD"
        Call Progreso("Generando archivo receta base...", 9, nCant): CreaArchivos "MRECETABASE"
        Call Progreso("Generando archivo receta base detalle...", 10, nCant): CreaArchivos "DRECETABASE"
        
        Call Progreso("Copiando Parametros Generales...", 11, nCant): CopiaArchivos "TPARAMETRO"
        Call Progreso("Copiando Archivo SubFamilias...", 12, nCant): CopiaArchivos "TSUBFAMILIA"
        Call Progreso("Copiando Archivo Tablas...", 13, nCant): CopiaArchivos "TTABLA"
        Call Progreso("Copiando Archivo Productos...", 14, nCant): CopiaArchivos "TPRODUCTO"
        Call Progreso("Copiando Archivo recetas venta...", 15, nCant): CopiaArchivos "MRECETAVENTA"
        Call Progreso("Copiando Archivo recetas venta detalle...", 16, nCant): CopiaArchivos "DRECETAVENTA"
        Call Progreso("Copiando Archivo recetas propiedad...", 17, nCant): CopiaArchivos "MRECETAPROPIEDAD"
        Call Progreso("Copiando Archivo recetas propiedad detalle...", 18, nCant): CopiaArchivos "DRECETAPROPIEDAD"
        Call Progreso("Copiando Archivo recetas base...", 19, nCant): CopiaArchivos "MRECETABASE"
        Call Progreso("Copiando Archivo recetas base detalle...", 20, nCant): CopiaArchivos "DRECETABASE"
        
        Sleep (3000) '3 segundos para q termine de realizar todo el proceso
        
        lblProgress.Caption = "Proceso terminado..."
    Else
        MuestraEvento 0
        MuestraEvento 1
        lblProgress.Visible = True
        pbProgress.Visible = True
        pbProgress.value = 100
        lblProgress.Caption = "Proceso terminado..."
    End If
    
    'Cambio: Rollback Jesus 13-07-2016
    msgError = ""
    lDescargo = True
    DscgFecSvrIni = FechaServidor()
    
    LlenaDatosInforest
    
    Dim CodPedido As String
    Dim RsCodPedido As Recordset
    
    Isql = "SELECT " & sTemporal & ".tCodigoPedido, " & sTemporal & ".fFecha FROM " & sTemporal & " Group By " & sTemporal & ".tCodigoPedido, " & sTemporal & ".fFecha Order By " & sTemporal & ".tCodigoPedido"
    Set RsCodPedido = Lib.OpenRecordset(Isql, Cn)
    
    'Proceso de validacion de cierre posterior
    Dim c As Integer
    If lKardexFechaIngreso Then
        If RsCodPedido.RecordCount > 0 Then
            RsCodPedido.MoveFirst
            
            MuestraEvento 0
            MuestraEvento 1
            lblProgress.Visible = True
            pbProgress.Visible = True
            pbProgress.value = 0
            cont = 0
            
            c = 0
            Do While Not RsCodPedido.EOF
                cont = cont + 1
                FechaTransaccion = Format(RsCodPedido!fFecha, "YYYY/mm/dd") & " 00:00:00"
                If lKardexFechaIngreso Then
                    If (ValidaCierrePosterior(FechaTransaccion, "", "", "95")) > 0 Then
                        Cn.Execute "Insert Into " & sPedidosSinDescargo & " values ('" & RsCodPedido!tCodigoPedido & "')"
                        c = c + 1
                    End If
                End If
                lblProgress.Caption = "Validando Registro " & cont & " de " & RsCodPedido.RecordCount
                pbProgress.value = (cont * 100) / RsCodPedido.RecordCount
                RsCodPedido.MoveNext
            Loop
            lblProgress.Caption = "Validación terminada..."
            If c > 0 Then
                If MsgBox("Existen cierres de inventario cuya fecha de cierre es posterior a la(s) fecha(s) de " & c & " Pedidos(s)." & vbCrLf & _
                            "Estos pedidos no se consideraran en el proceso de descargo." & vbCrLf & _
                            "¿Desea continuar con el Descargo de Ventas.?", vbExclamation + vbYesNo, "PROCESO DE DESCARGO: Alerta en Fecha de Pedidos") = vbNo Then Exit Sub
            End If
        End If
    End If
    If RsCodPedido.RecordCount > 0 Then
        RsCodPedido.MoveFirst
        
        MuestraEvento 2
        
        cmdOpcion.Enabled = False
        cmdSalir.Enabled = False
        
        lblProgreso.Visible = True
        pbProgreso.Visible = True
        pbProgreso.value = 0
        cont = 0
        cCorrect = 0
        cError = 0
               
        Do While Not RsCodPedido.EOF
            cont = cont + 1
            CodPedido = RsCodPedido!tCodigoPedido
            
            If Calcular("Select Count(*) as Codigo From " & sPedidosSinDescargo & " Where tCodigoPedido = '" & CodPedido & "'", Cn) > 0 Then
                cError = cError + 1
                Cn.Execute "update dpedido set ltransferido=NULL where tcodigopedido='" & CodPedido & "'"
                Cn.Execute "INSERT INTO " & sDtemporal & "(tEstado, tFecha, tPedido, tMensaje) VALUES(3, '" & RsCodPedido!fFecha & "','" & RsCodPedido!tCodigoPedido & "','Existe cierre posterior a pedido.')"
            Else
                ProcesoDescargo CodPedido, RsCodPedido!fFecha
                If lDescargoPedido = False Then
                    cError = cError + 1
                    'Cn.Execute "update dpedido set ltransferido=NULL where tcodigopedido='" & CodPedido & "'"
                    'Cn.Execute "INSERT INTO " & sDtemporal & "(tEstado, tFecha, tPedido, tMensaje) VALUES(1, '" & RsCodPedido!fFecha & "','" & RsCodPedido!tCodigoPedido & "','" & Replace(msgError, "'", "") & "')"
                Else
                    cCorrect = cCorrect + 1
                    'Cn.Execute "update dpedido set ltransferido=1 where tcodigopedido='" & CodPedido & "'"
                    'Cn.Execute "INSERT INTO " & sDtemporal & "(tEstado, tFecha, tPedido, tMensaje) VALUES(2, '" & RsCodPedido!fFecha & "','" & RsCodPedido!tCodigoPedido & "','Pedido descargo correctamente !!')"
                End If
            End If
            lblProgreso.Caption = "Procesando Registro " & cont & " de " & RsCodPedido.RecordCount
            pbProgreso.value = (cont * 100) / RsCodPedido.RecordCount
           
            RsCodPedido.MoveNext

        Loop
        
        RsCodPedido.MoveLast
        
        'Cambio: Guarda registro del descargo : Jesus 13-07-2016
        RegistraLogDescargo
        ' Fin cambio
        
        MuestraEvento 3
        
        cmdOpcion.Enabled = True
        cmdSalir.Enabled = True
        
        TraeTemporal
        
        aniVideo.AutoPlay = False
        aniVideo.Visible = False
        
        Screen.MousePointer = vbDefault
        
        MsgBox "Se realizaron " & Trim(str(RsCodPedido.RecordCount)) & " Transferencia(s). " & Trim(str(cCorrect)) & " con Exito", vbInformation, "Transferencia"
    Else
        aniVideo.AutoPlay = False
        aniVideo.Visible = False
        
        pbProgress.Visible = False
        pbProgress.value = 0
        lblProgress.Visible = False
        
        pbProgreso.Visible = False
        pbProgreso.value = 0
        lblProgreso.Visible = False
        
        lblProceso(0).ForeColor = &H808080
        imgProceso(0).Visible = False
        lblProceso(1).ForeColor = &H808080
        imgProceso(1).Visible = False
        lblProceso(2).ForeColor = &H808080
        imgProceso(2).Visible = False
        lblProceso(3).ForeColor = &H808080
        imgProceso(3).Visible = False
        
        Cn.Execute "Delete From " & sDtemporal
        TraeTemporal
        
        Screen.MousePointer = vbDefault
        MsgBox "No existe Ventas nuevas a transferir", vbExclamation, "Transferencia"
    End If
    
    Set RsTransferencia = Nothing
    If lCierre Then
        Isql = "SpUpdmCierre '" & Format(RsCierre!fCierre, "YYYY/mm/dd") & "', '" & Format(RsCierre!fRegistro, "YYYY/mm/dd HH:mm:ss") & "', '" & RsCierre!tUsuario & "', 1"
        CnAlmacen.Execute Isql
        Set RsCierre = Nothing
    End If
          
    Cn.Execute "drop table " & sTemporal
    
    'mdiAlmacen.StatusBar.Panels.Item(1).Text = "Local: " & LocalConectado
End Sub

Private Sub RegistraLogDescargo()
    Dim nCorrelativo As Integer
    Dim dwLen As Long
    Dim strString As String
    
    nCorrelativo = 0
    
    On Error GoTo ErrorLog
    
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    GetComputerName strString, dwLen
    strString = Left(strString, dwLen)
    
    EquipoUser = Get_User_Name
    EquipoName = strString
    EquipoIp = GetWanIP
    
    DscgFecIni = Format(dtpFecIni, "YYYY-mm-dd") & " " & Format(dtpHorIni, "HH:mm:ss")
    DscgFecFin = Format(dtpFecFin, "YYYY-mm-dd") & " " & Format(dtpHorFin, "HH:mm:ss")
    
    Set clsDescAlmacen = New clsAlmacen
    nCorrelativo = clsDescAlmacen.FunInsertaLogDescargo(nCorrelativo, Format(DscgFecSvrIni, "YYYY/mm/dd HH:mm:ss"), sUsuario, DscgFecIni, DscgFecFin, cont, cCorrect, cError, EquipoIp, EquipoName, EquipoUser)
    If nCorrelativo = -1 Then
        GoTo ErrorLog
    End If
    Exit Sub

ErrorLog:
    Screen.MousePointer = vbDefault
End Sub

Private Sub ProcesoDescargo(ByVal CodPedido As String, ByVal Fecha As Date)
    Dim cArea As String
    Dim c As Integer
    Dim RsPedidoDetalle As Recordset
    lDescargoPedido = True
    Isql = "SELECT " & sTemporal & ".tCodigoPedido, " & sTemporal & ".titem FROM " & sTemporal & " where tCodigoPedido = '" & CodPedido & "' Group By " & sTemporal & ".tCodigoPedido, " & sTemporal & ".titem Order By " & sTemporal & ".tCodigoPedido"
    Set RsPedidoDetalle = Lib.OpenRecordset(Isql, Cn)
    If RsPedidoDetalle.RecordCount = 0 Then: Exit Sub
    RsPedidoDetalle.MoveFirst
    Do While Not RsPedidoDetalle.EOF
        On Error GoTo ErrorDescargo
        'Cn.BeginTrans
        CnAlmacen.BeginTrans
        
        'Isql = "SELECT " & sTemporal & ".tCodigoPedido, fFecha, Plato, nCantidad, tItem, tDescargo, tEnlace, tTipoPedido, " & sTemporal & ".tCodigoProducto, nRecetaCantidad, tSubAreaAlm, tSubAreaInf, nFactor, nPrecioPromedio, " & sTemporal & ".lRecetaBase, " & sTemporal & ".lProducto," & sTemporal & ".tCodigoUnicoEtiqueta," & sTemporal & ".tDocumento, IsNull(mDocumento.tTipoDocumento,'') as tTipoDocumento, " & sTemporal & ".fDiaContable as fDiaContable FROM " & sTemporal & " INNER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO ON " & sTemporal & ".tCodigoProducto = " & sAlmacenMDB & ".dbo.TPRODUCTO.tCodigoProducto LEFT JOIN mDocumento on " & sTemporal & ".tDocumento = mDocumento.tDocumento where " & sTemporal & ".tCodigoPedido = " & CodPedido & " order by tCodigoPedido, tItem"
        Isql = " exec usp_inforest_DescargoVenta '" & sAlmacenMDB & "','" & Format(dtpFecIni.value, "yyyymmdd") & " " & Format(dtpHorIni.value, "HH:mm") & "','" & Format(dtpFecFin.value, "yyyymmdd") & " " & Format(dtpHorFin.value, "HH:mm") & "','" & sTemporal & "','" & RsPedidoDetalle!tItem & "','" & RsPedidoDetalle!tCodigoPedido & "',2"
        'Debug.Print Isql
        Set RsTransferencia = Lib.OpenRecordset(Isql, Cn)
        If RsTransferencia.RecordCount > 0 Then
            RsTransferencia.MoveFirst
            c = 0
            Do While Not RsTransferencia.EOF
                c = c + 1
                'Jesus 20-10-2016
                lDescargo = True
                If RsTransferencia!lRecetaBase And RsTransferencia!lProducto = False Then
                    DescargoInsumoRecetaBase RsTransferencia!tCodigoPedido, RsTransferencia!tCodigoProducto, (RsTransferencia!nRecetaCantidad / RsTransferencia!nFactor) * RsTransferencia!nCantidad, IIf(IsNull(RsTransferencia!fFecha), Date, RsTransferencia!fFecha), RsTransferencia!tCodigoUnicoEtiqueta, RsTransferencia!tDocumento, RsTransferencia!tTipoDocumento, RsTransferencia!fdiacontable
                Else
                    If Len(Trim(RsTransferencia!tSubAreaInf)) = 0 Then
                        cArea = RsTransferencia!tSubAreaAlm
                    Else
                        cArea = RsTransferencia!tSubAreaInf
                    End If
                    DescargoInsumo RsTransferencia!tCodigoPedido, cArea, RsTransferencia!tCodigoProducto, RsTransferencia!nCantidad, RsTransferencia!nRecetaCantidad, RsTransferencia!nFactor, RsTransferencia!tDescargo, IIf(IsNull(RsTransferencia!nPrecioPromedio), 0, RsTransferencia!nPrecioPromedio), IIf(IsNull(RsTransferencia!fFecha), Date, RsTransferencia!fFecha), RsTransferencia!tCodigoUnicoEtiqueta, RsTransferencia!tDocumento, RsTransferencia!tTipoDocumento, RsTransferencia!fdiacontable
                End If
                If lDescargo = False Then
                    'Cn.Execute "update dpedido set ltransferido= NULL where tcodigopedido='" & RsPedidoDetalle!tCodigoPedido & "' and  titem='" & RsPedidoDetalle!tItem & "'"
                    'GoTo ErrorDescargo
                    Exit Do
                End If
                RsTransferencia.MoveNext
            Loop
            
            RsTransferencia.MoveLast
        End If
        
        If lDescargo Then
            'Cn.CommitTrans
            CnAlmacen.CommitTrans
            Cn.Execute "update dpedido set ltransferido=1 where tcodigopedido='" & RsPedidoDetalle!tCodigoPedido & "' and  titem='" & RsPedidoDetalle!tItem & "'"
            Cn.Execute "INSERT INTO " & sDtemporal & "(tEstado, tFecha, tPedido, tItem , tMensaje) VALUES(2, '" & Fecha & "','" & RsPedidoDetalle!tCodigoPedido & "','" & RsPedidoDetalle!tItem & "','Descargo correctamente !!')"
        Else
            'Cn.RollbackTrans
            CnAlmacen.RollbackTrans
            Cn.Execute "update dpedido set ltransferido= NULL where tcodigopedido='" & RsPedidoDetalle!tCodigoPedido & "' and  titem='" & RsPedidoDetalle!tItem & "'"
            Cn.Execute "INSERT INTO " & sDtemporal & "(tEstado, tFecha,  tPedido, tItem, tMensaje) VALUES(1, '" & Fecha & "','" & RsPedidoDetalle!tCodigoPedido & "','" & RsPedidoDetalle!tItem & "','" & Replace(msgError, "'", "") & "')"
            If lDescargoPedido Then
                lDescargoPedido = False
            End If
        End If

        RsPedidoDetalle.MoveNext
    Loop
    Exit Sub
    
ErrorDescargo:
    'Cn.RollbackTrans
    CnAlmacen.RollbackTrans
    lDescargo = False
    lDescargoPedido = False
    MsgBox error, vbInformation, sMensaje
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.value > dtpFecFin.value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub Form_Load()
    Centrar Me
    
    Screen.MousePointer = vbHourglass
    
    sDtemporal = dbTemporal(sUsuario, 5, "tEstado", "int", _
                                         "tPedido", "nVarChar(100)", _
                                         "tItem", "nVarChar(100)", _
                                         "tFecha", "nVarchar(30)", _
                                         "tmensaje", "nVarchar(1000)")
    TraeTemporal
    
      
    dtpFecIni.value = Date
    dtpFecFin.value = Date
    
    Screen.MousePointer = vbDefault
    
On Error Resume Next

    aniVideo.Open App.Path & "\Bmps\FileMove.avi"

On Error GoTo ErrorParametroAlmacen
    
    lKardexFechaIngreso = Calcular("select IsNull(lKardexFechaIngreso,0) as Codigo from tParametro", IIf(verificaAlmacenRemoto = "ON", CnAlmacenRemoto, CnAlmacen))
    Exit Sub

ErrorParametroAlmacen:
    MsgBox "Parametros incompletos. Verifique la configuración en los parametros del Almacen.", vbExclamation, ""
    Frame1.Enabled = False
    cmdOpcion.Enabled = False
    cmdSalir.Enabled = False
    Screen.MousePointer = vbDefault
End Sub

Public Sub DescargoInsumo(Documento As String, Area As String, Insumo As String, Cantidad As Double, CantidadReceta As Double, Factor As Double, Tipo As String, PrecioPromedio As Double, Fecha As String, tEtiqueta As String, tDocumentoVenta As String, tTipoDocumentoVenta As String, fdiacontable As String)
    Dim nStock    As Double
    Dim nCorrela  As Double
    Dim nCantidad As Double
    Dim sCodProd  As String
      
    Screen.MousePointer = vbHourglass
    If IsNull(Area) Or Len(Area) < 1 Then
        Exit Sub
    End If
   
    If IsNull(Insumo) Or Len(Insumo) < 1 Then
        Exit Sub
    End If
    nCantidad = (Cantidad * CantidadReceta) / IIf(Tipo = "D", 1, Factor)
    'Descargo hacia el Almacen Central
    On Error GoTo ErrorAlmacen
    'Cn.BeginTrans
    Set clsDescAlmacen = New clsAlmacen
    If Area = "000" Then
        If Calcular("select top 1 isnull(lstockdescargo,0) as codigo from tparametro", Cn) Then
            If Calcular("select isnull(nstockactual,0) as Codigo from tproducto where tcodigoproducto='" & Insumo & "'", CnAlmacen) < nCantidad Then: msgError = "Insumo insuficiente para Descargo Almacen Central, Insumo: " & Insumo & "": GoTo ErrorAlmacen
        End If
        sCodProd = clsDescAlmacen.FunInsertamKardex(Insumo, 0, "95", True, Area, Documento, 0, PrecioPromedio * nCantidad, nCantidad, 0, 0, "01", PrecioPromedio, PrecioPromedio, Format(IIf(lCierre, Fecha, fFechaCierre), "YYYY/mm/dd HH:mm:ss"), sUsuario, "")
        If sCodProd = "" Then
            GoTo ErrorAlmacen
        End If
    Else
        If Calcular("select top 1 isnull(lstockdescargo,0) as codigo from tparametro", Cn) Then
            If Calcular("select isnull(nstockactual,0) as Codigo from tsubstock where tcodigoproducto='" & Insumo & "' and tCodigoSubArea='" & Area & "'", CnAlmacen) < nCantidad Then: msgError = "Sin Stock, Area: " & Area & ", Insumo: " & Insumo & "": GoTo ErrorAlmacen
        End If
        sCodProd = clsDescAlmacen.FunInsertamSubKardex(Area, Insumo, 0, "", Format(IIf(lCierre, Fecha, fFechaCierre), "YYYY/mm/dd HH:mm:ss"), sUsuario, "95", True, Documento, 0, nCantidad, PrecioPromedio * nCantidad, 0, 0, "01", tEtiqueta, tDocumentoVenta, tTipoDocumentoVenta, Format(fdiacontable, "YYYY/mm/dd HH:mm:ss"), Fecha)
        If sCodProd = "" Then
            GoTo ErrorAlmacen
        End If
    End If
    Set clsDescAlmacen = Nothing
    'Cn.CommitTrans
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorAlmacen:
    'Cn.RollbackTrans
    lDescargo = False
    msgError = msgError + " " + err.Description
    'MsgBox "Ocurrio un error al momento de descargar el Pedido " & Trim(Documento), vbCritical, "Transferencia"
    'mdiAdministracion.StatusBar.Panels.Item(2).Text = "Ocurrio un error al momento de descargar el Pedido " & Trim(Documento)
    Screen.MousePointer = vbDefault
End Sub

Sub DescargoInsumoRecetaBase(ByVal diCodigoPedido As String, ByVal diCodigoProducto As String, ByVal diCantidad As Double, ByVal diFecha As String, tEtiqueta As String, tDocumentoVenta As String, tTipoDocumentoVenta As String, fdiacontable As String)
    Dim RsDescargoReceta As ADODB.Recordset
    
    Isql = "SELECT RB.tRecetaBase, RB.tSubArea, RB.tCodigoProducto, isnull(PR.lRecetaBase, 0) as lRecetaBase, isnull(RB.lProducto, 0) as lProducto, isnull(RB.nCantidad, 0) as nCantidadReceta, isnull(PR.nFactor, 0) as nFactor, ISNULL(PR.nPrecioPromedio, 0) AS nPrecioPromedio " & _
           "FROM " & sAlmacenMDB & ".dbo.DRECETABASE RB INNER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO PR ON RB.tCodigoProducto = PR.tCodigoProducto INNER JOIN " & sAlmacenMDB & ".dbo.MRECETABASE MRB ON RB.tRecetaBase = MRB.tRecetaBase " & _
           "WHERE MRB.tCodigoProducto='" & diCodigoProducto & "' AND (RB.lDescargo=1)"

    Set RsDescargoReceta = Lib.OpenRecordset(Isql, Cn)
    If RsDescargoReceta.RecordCount > 0 Then
       RsDescargoReceta.MoveFirst
       Do While Not RsDescargoReceta.EOF
          If RsDescargoReceta!lRecetaBase And RsDescargoReceta!lProducto = False Then
             DescargoInsumoRecetaBase Trim(Left(Left(diCodigoPedido, 10) & Trim(RsDescargoReceta!tRecetaBase), 15)), RsDescargoReceta!tCodigoProducto, (RsDescargoReceta!nCantidadReceta / RsDescargoReceta!nFactor) * diCantidad, diFecha, tEtiqueta, tDocumentoVenta, tTipoDocumentoVenta, fdiacontable
          Else
             DescargoInsumo Trim(Left(Left(diCodigoPedido, 10) & Trim(RsDescargoReceta!tRecetaBase), 15)), RsDescargoReceta!tSubArea, RsDescargoReceta!tCodigoProducto, diCantidad, RsDescargoReceta!nCantidadReceta, RsDescargoReceta!nFactor, "R", RsDescargoReceta!nPrecioPromedio, diFecha, tEtiqueta, tDocumentoVenta, tTipoDocumentoVenta, fdiacontable
          End If
          RsDescargoReceta.MoveNext
       Loop
    End If
    Set RsDescargoReceta = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cn.Execute "Delete From " & sDtemporal
End Sub

